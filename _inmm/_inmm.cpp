#include <stdio.h>
#include <Windows.h>
#include <WinGDI.h>
#include <ddraw.h>
#include <MMSystem.h>

#define _INMM_LOG_OUTPUT _DEBUG
//#define _INMM_PERF_LOG

static HDC hDesktopDC = NULL;
static HPALETTE hPalette = NULL;

// フォントテーブル
const int nMaxFontTable = 32;
typedef struct
{
	HFONT hFont;
	char szFaceName[LF_FACESIZE];
	int nHeight;
	int nBold;
	int nAdjustX;
	int nAdjustY;
} FONT_TABLE;
static FONT_TABLE fontTable[nMaxFontTable];

// プリセットカラーテーブル
const int nMaxColorTable = 'Z' - 'A' + 1;
typedef struct
{
	BYTE r;
	BYTE g;
	BYTE b;
} COLOR_TABLE;
static COLOR_TABLE colorTable[nMaxColorTable] =
{
	{   0,  64,   0 }, // A
	{   0,   0, 192 }, // B
	{ 128,   0,   0 }, // C
	{ 255, 255,   0 }, // D
	{ 255,   0, 255 }, // E
	{ 255, 128,   0 }, // F
	{   0, 128,   0 }, // G
	{ 128, 128, 128 }, // H
	{   0,   0, 255 }, // I
	{   0, 192,   0 }, // J
	{   0,   0,   0 }, // K
	{   0,   0,   0 }, // L
	{ 100, 255, 100 }, // M
	{ 255, 100, 100 }, // N
	{   0,   0,   0 }, // O
	{   0,   0,   0 }, // P
	{   0,   0,   0 }, // Q
	{ 240,   0,   0 }, // R
	{   0,   0,   0 }, // S
	{   0,   0,   0 }, // T
	{   0,   0,   0 }, // U
	{   0,   0,   0 }, // V
	{ 255, 255, 255 }, // W
	{   0,   0,   0 }, // X
	{ 208, 232, 248 }, // Y
	{   0,   0,   0 }, // Z
};

// コイン画像表示用のメモリデバイスコンテキストのハンドル
static HDC hDCCoin = NULL;

// コイン画像のビットマップハンドル
static HBITMAP hBitmapCoin = NULL;

// コイン画像読み込み前のビットマップハンドル退避先
static HBITMAP hBitmapCoinPrev = NULL;

// コイン画像のファイル名
static char szCoinFileName[MAX_PATH] = "Gfx\\Fonts\\Modern\\Additional\\coin.bmp";

// コイン画像の表示位置調整
static int nCoinAdjustX = 0;
static int nCoinAdjustY = 0;

// コイン画像の大きさ
static const int nCoinWidth = 12;
static const int nCoinHeight = 12;

// 最大表示可能行数
static int nMaxLines = 84;

// 単語内の最大表示可能文字数
static int nMaxWordChars = 24;

// WINMM.DLLのモジュールハンドル
static HMODULE hWinMmDll = NULL;

// 文字列処理バッファ
#define STRING_BUFFER_SIZE 1024

// ログ出力用ファイルポインタ
#if _INMM_LOG_OUTPUT
static FILE *fpLog = NULL;
#endif

// 公開関数
int WINAPI GetTextWidth(LPCBYTE lpString, int nMagicCode, DWORD dwFlags);
void WINAPI TextOutDC0(int x, int y, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags);
void WINAPI TextOutDC1(LPRECT lpRect, int x, int y, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags);
void WINAPI TextOutDC2(LPRECT lpRect, int *px, int *py, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags);
int WINAPI CalcLineBreak(LPBYTE lpBuffer, LPCBYTE lpString);
int WINAPI strnlen0(LPCBYTE lpString, int nMax);

// 内部関数
static int GetTextWidthWord(LPCBYTE lpString, int nLen, int nMagicCode);
static void TextOutWord(LPCBYTE lpString, int nLen, LPPOINT lpPoint, LPRECT lpRect, HDC hDC, int nMagicCode, COLORREF color, DWORD dwFlags);
static void ShowCoinImage(LPPOINT lpPoint, HDC hDC);
static int CalcColorWordWrap(LPBYTE lpBuffer, LPCBYTE lpString);
static int CalcNumberWordWrap(LPBYTE lpBuffer, LPCBYTE lpString);
static void Init();
static void LoadIniFile();
static void InitFont();
static void InitPalette();
static void InitCoinImage();
static void Terminate();
static bool ReadLine(FILE *fp, char *buf, int len);

// winmm.dllのAPI転送用
typedef MCIERROR (WINAPI *LPFNMCISENDCOMMANDA)(MCIDEVICEID, UINT, DWORD, DWORD);
typedef MMRESULT (WINAPI *LPFNTIMEBEGINPERIOD)(UINT);
typedef MMRESULT (WINAPI *LPFNTIMEGETDEVCAPS)(LPTIMECAPS, UINT);
typedef DWORD (WINAPI *LPFNTIMEGETTIME)(VOID);
typedef MMRESULT (WINAPI *LPFNTIMEKILLEVENT)(UINT);
typedef MMRESULT (WINAPI *LPFNTIMESETEVENT)(UINT, UINT, LPTIMECALLBACK, DWORD, UINT);

static LPFNMCISENDCOMMANDA lpfnMciSendCommandA = NULL;
static LPFNTIMEBEGINPERIOD lpfnTimeBeginPeriod = NULL;
static LPFNTIMEGETDEVCAPS lpfnTimeGetDevCaps = NULL;
static LPFNTIMEGETTIME lpfnTimeGetTime = NULL;
static LPFNTIMEKILLEVENT lpfnTimeKillEvent = NULL;
static LPFNTIMESETEVENT lpfnTimeSetEvent = NULL;

MCIERROR WINAPI _mciSendCommandA(MCIDEVICEID IDDevice, UINT uMsg, DWORD fdwCommand, DWORD dwParam);
MMRESULT WINAPI _timeBeginPeriod(UINT uPeriod);
MMRESULT WINAPI _timeGetDevCaps(LPTIMECAPS ptc, UINT cbtc);
DWORD WINAPI _timeGetTime(VOID);
MMRESULT WINAPI _timeKillEvent(UINT uTimerID);
MMRESULT WINAPI _timeSetEvent(UINT uDelay, UINT uResolution, LPTIMECALLBACK lpTimeProc, DWORD dwUser, UINT fuEvent);

//
// DLLメイン関数
//
// パラメータ
//   hInstDLL: DLL モジュールのハンドル
//   fdwReason: 関数を呼び出す理由
//   lpvReserved: 予約済み
// 戻り値
//   処理が成功すればTRUEを返す
//   
BOOL WINAPI DllMain(HINSTANCE hInstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	switch (fdwReason)
	{
	case DLL_PROCESS_ATTACH:
		Init();
		break;
	case DLL_THREAD_ATTACH:
		break;
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		Terminate();
		break;
	default:
		break;
	}
	return TRUE;
}

//
// 文字列幅を取得する (オリジナル版互換)
//
// パラメータ
//   lpString		対象文字列
//   nMagicCode		フォント指定用マジックコード
//   dwFlags		フラグ(1で太文字/2で影付き)
// 戻り値
//   文字列幅
//
int WINAPI GetTextWidth(LPCBYTE lpString, int nMagicCode, DWORD dwFlags)
{
#if _INMM_LOG_OUTPUT
#ifdef _INMM_PERF_LOG
	LARGE_INTEGER nFreq;
	LARGE_INTEGER nBefore;
	LARGE_INTEGER nAfter;
	memset(&nFreq, 0, sizeof(nFreq));
	memset(&nBefore, 0, sizeof(nBefore));
	memset(&nAfter, 0, sizeof(nAfter));
	QueryPerformanceFrequency(&nFreq);
	QueryPerformanceCounter(&nBefore);
#endif
#endif

	LPCBYTE s = lpString;
	LPCBYTE p = s;
	int nWidth = 0;
	int nMaxWidth = 0;
	int nLines = 0;

	if (fontTable[nMagicCode].hFont == NULL)
	{
		nMagicCode = 0;
	}
	HFONT hFontOld = (HFONT)SelectObject(hDesktopDC, fontTable[nMagicCode].hFont);

	while (*s != '\0')
	{
		switch (*s)
		{
		case 0x81:
			// §
			if (*(s + 1) == 0x98)
			{
				// §の後に英大文字が続けばプリセットの色指定
				if ((*(s + 2) >= 'A') && (*(s + 2) <= 'Z'))
				{
					if (s > p)
					{
						nWidth += GetTextWidthWord(p, s - p, nMagicCode);
					}
					s += 3;
					p = s;
					continue;
				}
				// §§ならば元の色に戻す
				if ((*(s + 2) == 0x81) && (*(s + 3) == 0x98))
				{
					if (s > p)
					{
						nWidth += GetTextWidthWord(p, s - p, nMagicCode);
					}
					s += 4;
					p = s;
					continue;
				}
				// 次の3文字が0～7ならば直値の色指定
				if ((*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7') && (*(s + 4) >= '0') && (*(s + 4) <= '7'))
				{
					if (s > p)
					{
						nWidth += GetTextWidthWord(p, s - p, nMagicCode);
					}
					s += 5;
					p = s;
					continue;
				}
				s += 2;
				continue;
			}
			// Shift_JISの2バイト目
			if (((*(s + 1) >= 0x40) && (*(s + 1) <= 0x7E)) || ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xFC)))
			{
				s += 2;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0xA7: // ｧ
			// ｧの後に英大文字が続けばプリセットの色指定
			if ((*(s + 1) >= 'A') && (*(s + 1) <= 'Z'))
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				s += 2;
				p = s;
				continue;
			}
			// ｧｧならば元の色に戻す
			if (*(s + 1) == 0xA7)
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				s += 2;
				p = s;
				continue;
			}
			// 次の3文字が0～7ならば直値の色指定
			if ((*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7'))
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				s += 4;
				p = s;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case '%':
			// 次の3文字が0～7ならば直値の色指定
			if ((*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7'))
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				s += 4;
				p = s;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case '$':
			// コイン画像があれば幅を加算する
			if (hBitmapCoin != NULL)
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				nWidth += nCoinWidth;
				s++;
				p = s;
				continue;
			}
			// コイン画像がなければ単体の文字として扱う
			s++;
			continue;

		case '\\': // エスケープ記号
			// エスケープ記号の後に%/ｧ/$が続けば単体の文字として扱う
			if (*(s + 1) == '%' || *(s + 1) == 0xA7 || *(s + 1) == '$')
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				s++;
				nWidth += GetTextWidthWord(s, 1, nMagicCode);
				s++;
				p = s;
				continue;
			}
			// エスケープ記号の後に§が続けば単体の文字として扱う
			if ((*(s + 1) == 0x81) && (*(s + 2) == 0x98))
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				s++;
				nWidth += GetTextWidthWord(s, 2, nMagicCode);
				s += 2;
				p = s;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0x0A:
		case 0x0D:
			// 改行コード
			if (s > p)
			{
				nWidth += GetTextWidthWord(p, s - p, nMagicCode);
			}
			s++;
			p = s;
			if (nWidth > nMaxWidth)
			{
				nMaxWidth = nWidth;
				nWidth = 0;
			}
			nLines++;
			// 最大表示可能行数に到達していなければ計算を継続する
			if (nLines < nMaxLines)
			{
				continue;
			}
			// 最大表示可能行数に到達したら計算を打ち切る
			*(char *)(s - 1) = '\0';
			break;

		case 0x01:
		case 0x02:
		case 0x03:
		case 0x04:
		case 0x05:
		case 0x06:
		case 0x07:
		case 0x08:
		case 0x09:
		case 0x0B:
		case 0x0C:
		case 0x0E:
		case 0x0F:
		case 0x10:
		case 0x11:
		case 0x12:
		case 0x13:
		case 0x14:
		case 0x15:
		case 0x16:
		case 0x17:
		case 0x18:
		case 0x19:
		case 0x1A:
		case 0x1B:
		case 0x1C:
		case 0x1D:
		case 0x1E:
		case 0x1F:
			// 改行以外の制御コードならば読み飛ばす
			if (s > p)
			{
				nWidth += GetTextWidthWord(p, s - p, nMagicCode);
			}
			s++;
			p = s;
			continue;

		case 0x82:
		case 0x83:
		case 0x84:
		case 0x85:
		case 0x86:
		case 0x87:
		case 0x88:
		case 0x89:
		case 0x8A:
		case 0x8B:
		case 0x8C:
		case 0x8D:
		case 0x8E:
		case 0x8F:
		case 0x90:
		case 0x91:
		case 0x92:
		case 0x93:
		case 0x94:
		case 0x95:
		case 0x96:
		case 0x97:
		case 0x98:
		case 0x99:
		case 0x9A:
		case 0x9B:
		case 0x9C:
		case 0x9D:
		case 0x9E:
		case 0x9F:
		case 0xE0:
		case 0xE1:
		case 0xE2:
		case 0xE3:
		case 0xE4:
		case 0xE5:
		case 0xE6:
		case 0xE7:
		case 0xE8:
		case 0xE9:
		case 0xEA:
		case 0xEB:
		case 0xEC:
		case 0xED:
		case 0xEE:
		case 0xEF:
		case 0xF0:
		case 0xF1:
		case 0xF2:
		case 0xF3:
		case 0xF4:
		case 0xF5:
		case 0xF6:
		case 0xF7:
		case 0xF8:
		case 0xF9:
		case 0xFA:
		case 0xFB:
		case 0xFC:
			// Shift_JISの2バイト目
			if (((*(s + 1) >= 0x40) && (*(s + 1) <= 0x7E)) || ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xFC)))
			{
				s += 2;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		default:
			// それ以外ならば単体の文字として扱う
			s++;
			continue;
		}
		// 最大表示可能行数に到達した時に計算を打ち切るためのbreak
		// 続ける場合はswitch文の中にcontinueを書いてループの先頭へ飛ぶこと
		break;
	}
	if (s > p)
	{
		nWidth += GetTextWidthWord(p, s - p, nMagicCode);
	}
	if (nWidth > nMaxWidth)
	{
		nMaxWidth = nWidth;
	}
	if (nMaxWidth > 0)
	{
		nMaxWidth += ((dwFlags & 2) != 0) ? 1 : 0;
	}

	SelectObject(hDesktopDC, hFontOld);

#if _INMM_LOG_OUTPUT
#ifdef _INMM_PERF_LOG
	QueryPerformanceCounter(&nAfter);
	fprintf(fpLog, "[PERF] GetTextWidth: %lf\n", (double)(nAfter.QuadPart - nBefore.QuadPart) * 1000 / nFreq.QuadPart);
#endif
	fprintf(fpLog, "GetTextWidth: %s (%d) %X => %d\n", lpString, nMagicCode, dwFlags, nMaxWidth);
#endif

	return nMaxWidth;
}

//
// 単語の文字列幅を取得する
//
// パラメータ
//   lpString		対象文字列
//   nLen           文字列長(Byte単位)
//   nMagicCode		フォント指定用マジックコード
// 戻り値
//   文字列幅
//
int GetTextWidthWord(LPCBYTE lpString, int nLen, int nMagicCode)
{
	SIZE size;
	GetTextExtentPoint32(hDesktopDC, (LPCSTR)lpString, nLen, &size);

#if _INMM_LOG_OUTPUT
#ifndef _INMM_PERF_LOG
	char buf[STRING_BUFFER_SIZE];
	lstrcpyn(buf, (LPCSTR)lpString, STRING_BUFFER_SIZE);
	buf[nLen] = '\0';
	fprintf(fpLog, "  GetTextWidthWord %s %d %d => %d\n", buf, nLen, nMagicCode, size.cx);
#endif
#endif

	return size.cx;
}

//
// 文字列を出力する (オリジナル版互換)
//
// パラメータ
//   x				出力先X座標
//   y				出力先Y座標
//   lpString		対象文字列
//   lpDDS			DirectDrawサーフェースへのポインタ
//   nMagicCode		フォント指定用マジックコード
//   dwColor		色指定(R-G-Bを5bit-6bit-5bitのWORD値で)
//   dwFlags		フラグ(1で太文字/2で影付き)
//
void WINAPI TextOutDC0(int x, int y, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags)
{
	TextOutDC2(NULL, &x, &y, lpString, lpDDS, nMagicCode, dwColor, dwFlags);
}

//
// 文字列を出力する (クリッピング対応版)
//
// パラメータ
//   lpRect         クリッピング矩形(NULLの場合はクリッピングなし)
//   x				出力先X座標
//   y				出力先Y座標
//   lpString		対象文字列
//   lpDDS			DirectDrawサーフェースへのポインタ
//   nMagicCode		フォント指定用マジックコード
//   dwColor		色指定(R-G-Bを5bit-6bit-5bitのWORD値で)
//   dwFlags		フラグ(1で太文字/2で影付き)
//

void WINAPI TextOutDC1(LPRECT lpRect, int x, int y, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags)
{
	TextOutDC2(lpRect, &x, &y, lpString, lpDDS, nMagicCode, dwColor, dwFlags);
}

//
// 文字列を出力する (出力位置更新対応版)
//
// パラメータ
//   lpRect         クリッピング矩形(NULLの場合はクリッピングなし)
//   px				出力先X座標への参照
//   py				出力先Y座標への参照
//   lpString		対象文字列
//   lpDDS			DirectDrawサーフェースへのポインタ
//   nMagicCode		フォント指定用マジックコード
//   dwColor		色指定(R-G-Bを5bit-6bit-5bitのWORD値で)
//   dwFlags		フラグ(1で太文字/2で影付き)
//
void WINAPI TextOutDC2(LPRECT lpRect, int *px, int *py, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags)
{
#if _INMM_LOG_OUTPUT
	fprintf(fpLog, "TextOutDC2 %s (%d,%d) %X %d %X %X\n", lpString, *px, *py, lpDDS, nMagicCode, dwColor, dwFlags);
#ifdef _INMM_PERF_LOG
	LARGE_INTEGER nFreq;
	LARGE_INTEGER nBefore;
	LARGE_INTEGER nAfter;
	memset(&nFreq, 0, sizeof(nFreq));
	memset(&nBefore, 0, sizeof(nBefore));
	memset(&nAfter, 0, sizeof(nAfter));
	QueryPerformanceFrequency(&nFreq);
	QueryPerformanceCounter(&nBefore);
#endif
#endif

	LPCBYTE s = lpString;
	LPCBYTE p = s;
	COLORREF color = RGB(((dwColor & 0x0000F800) >> 11) << 3, ((dwColor & 0x000007E0) >> 5) << 2, (dwColor & 0x0000001F) << 3);
	COLORREF colorOld = color;

	HDC hDC;
	lpDDS->GetDC(&hDC);
	SetBkMode(hDC, TRANSPARENT);
	if (fontTable[nMagicCode].hFont == NULL)
	{
		nMagicCode = 0;
	}
	HFONT hFontOld = (HFONT)SelectObject(hDC, fontTable[nMagicCode].hFont);

	POINT pt = { *px, *py };

	while (*s != '\0')
	{
		switch (*s)
		{
		case 0x81:
			// §
			if (*(s + 1) == 0x98)
			{
				// §の後に英大文字が続けばプリセットの色指定
				if ((*(s + 2) >= 'A') && (*(s + 2) <= 'Z'))
				{
					// 出力する文字列があれば表示する
					if (s > p)
					{
						TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
					}
					colorOld = color;
					color = RGB(colorTable[*(s + 2) - 'A'].r, colorTable[*(s + 2) - 'A'].g, colorTable[*(s + 2) - 'A'].b);
					s += 3;
					p = s;
					continue;
				}
				// §§ならば元の色に戻す
				if ((*(s + 2) == 0x81) && (*(s + 3) == 0x98))
				{
					// 出力する文字列があれば表示する
					if (s > p)
					{
						TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
					}
					COLORREF temp = color;
					color = colorOld;
					colorOld = temp;
					s += 4;
					p = s;
					continue;
				}
				// 次の3文字が0～7ならば直値の色指定
				if ((*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7') && (*(s + 4) >= '0') && (*(s + 4) <= '7'))
				{
					// 出力する文字列があれば表示する
					if (s > p)
					{
						TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
					}
					colorOld = color;
					color = RGB((*(s + 4) - '0') << 5, (*(s + 3) - '0') << 5, (*(s + 2) - '0') << 5);
					s += 5;
					p = s;
					continue;
				}
				s += 2;
				continue;
			}
			// Shift_JISの2バイト目
			if (((*(s + 1) >= 0x40) && (*(s + 1) <= 0x7E)) || ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xFC)))
			{
				s += 2;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0xA7: // ｧ
			// ｧの後に英大文字が続けばプリセットの色指定
			if ((*(s + 1) >= 'A') && (*(s + 1) <= 'Z'))
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				colorOld = color;
				color = RGB(colorTable[*(s + 1) - 'A'].r, colorTable[*(s + 1) - 'A'].g, colorTable[*(s + 1) - 'A'].b);
				s += 2;
				p = s;
				continue;
			}
			// ｧｧならば元の色に戻す
			if (*(s + 1) == 0xA7)
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				COLORREF temp = color;
				color = colorOld;
				colorOld = temp;
				s += 2;
				p = s;
				continue;
			}
			// 次の3文字が0～7ならば直値の色指定
			if ((*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7'))
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				colorOld = color;
				color = RGB((*(s + 3) - '0') << 5, (*(s + 2) - '0') << 5, (*(s + 1) - '0') << 5);
				s += 4;
				p = s;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case '%':
			// 次の3文字が0～7ならば直値の色指定
			if ((*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7'))
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				colorOld = color;
				color = RGB((*(s + 3) - '0') << 5, (*(s + 2) - '0') << 5, (*(s + 1) - '0') << 5);
				s += 4;
				p = s;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case '$':
			// コイン画像があれば表示する
			if (hBitmapCoin != NULL)
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				// コイン画像を表示する
				ShowCoinImage(&pt, hDC);
				s++;
				p = s;
				continue;
			}
			// コイン画像がなければ単体の文字として扱う
			s++;
			continue;

		case '\\': // エスケープ記号
			// エスケープ記号の後に%/ｧ/$が続けば単体の文字として扱う
			if (*(s + 1) == '%' || *(s + 1) == 0xA7 || *(s + 1) == '$')
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				s++;
				TextOutWord(s, 1, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				s++;
				p = s;
				continue;
			}
			// エスケープ記号の後に§が続けば単体の文字として扱う
			if ((*(s + 1) == 0x81) && (*(s + 2) == 0x98))
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				s++;
				TextOutWord(s, 2, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				s += 2;
				p = s;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0x01:
		case 0x02:
		case 0x03:
		case 0x04:
		case 0x05:
		case 0x06:
		case 0x07:
		case 0x08:
		case 0x09:
		case 0x0B:
		case 0x0C:
		case 0x0E:
		case 0x0F:
		case 0x10:
		case 0x11:
		case 0x12:
		case 0x13:
		case 0x14:
		case 0x15:
		case 0x16:
		case 0x17:
		case 0x18:
		case 0x19:
		case 0x1A:
		case 0x1B:
		case 0x1C:
		case 0x1D:
		case 0x1E:
		case 0x1F:
			// 出力する文字列があれば表示する
			if (s > p)
			{
				TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
			}
			// 制御コードを読み飛ばす
			s++;
			p = s;
			continue;

		case 0x82:
		case 0x83:
		case 0x84:
		case 0x85:
		case 0x86:
		case 0x87:
		case 0x88:
		case 0x89:
		case 0x8A:
		case 0x8B:
		case 0x8C:
		case 0x8D:
		case 0x8E:
		case 0x8F:
		case 0x90:
		case 0x91:
		case 0x92:
		case 0x93:
		case 0x94:
		case 0x95:
		case 0x96:
		case 0x97:
		case 0x98:
		case 0x99:
		case 0x9A:
		case 0x9B:
		case 0x9C:
		case 0x9D:
		case 0x9E:
		case 0x9F:
		case 0xE0:
		case 0xE1:
		case 0xE2:
		case 0xE3:
		case 0xE4:
		case 0xE5:
		case 0xE6:
		case 0xE7:
		case 0xE8:
		case 0xE9:
		case 0xEA:
		case 0xEB:
		case 0xEC:
		case 0xED:
		case 0xEE:
		case 0xEF:
		case 0xF0:
		case 0xF1:
		case 0xF2:
		case 0xF3:
		case 0xF4:
		case 0xF5:
		case 0xF6:
		case 0xF7:
		case 0xF8:
		case 0xF9:
		case 0xFA:
		case 0xFB:
		case 0xFC:
			// Shift_JISの2バイト目
			if (((*(s + 1) >= 0x40) && (*(s + 1) <= 0x7E)) || ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xFC)))
			{
				s += 2;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		default:
			// それ以外ならば単体の文字として扱う
			s++;
			continue;
		}
	}
	// 出力する文字列があれば表示する
	if (s > p)
	{
		TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
	}

	SelectObject(hDC, hFontOld);
	lpDDS->ReleaseDC(hDC);

	*px = pt.x;
	*py = pt.y;

#if _INMM_LOG_OUTPUT
#ifdef _INMM_PERF_LOG
	QueryPerformanceCounter(&nAfter);
	fprintf(fpLog, "[PERF] TextOutDC1: %lf\n", (double)(nAfter.QuadPart - nBefore.QuadPart) * 1000 / nFreq.QuadPart);
#endif
#endif
}

//
// 単語を出力する(途中で色変更はなし)
//
// パラメータ
//   lpString		対象文字列
//   nLen           文字列長(Byte単位)
//   lpPoint        出力先座標
//   lpRect         クリッピング矩形(NULLの場合はクリッピングなし)
//   hDC			デバイスコンテキストのハンドル
//   color			色指定
//   dwFlags		フラグ(1で太文字/2で影付き)
// 
void TextOutWord(LPCBYTE lpString, int nLen, LPPOINT lpPoint, LPRECT lpRect, HDC hDC, int nMagicCode, COLORREF color, DWORD dwFlags)
{
#if _INMM_LOG_OUTPUT
#ifndef _INMM_PERF_LOG
	fprintf(fpLog, "  TextOutWord %s %d (%d,%d) (%d,%d)-(%d,%d) %X %d %X %X\n", lpString, nLen, lpPoint->x, lpPoint->y,
		(lpRect != NULL) ? lpRect->left : 0, (lpRect != NULL) ? lpRect->top : 0, (lpRect != NULL) ? lpRect->right : 0, (lpRect != NULL) ? lpRect->bottom : 0,
		hDC, nMagicCode, (DWORD)color, dwFlags);
#endif
#endif

	int x = lpPoint->x + fontTable[nMagicCode].nAdjustX;
	int y = lpPoint->y + fontTable[nMagicCode].nAdjustY;

	// 影を付ける場合は出力位置の右下に黒で出力する
	bool bShadowed = ((dwFlags & 2) != 0);
	if (bShadowed)
	{
		SetTextColor(hDC, RGB(0, 0, 0));
		if (lpRect != NULL)
		{
			ExtTextOut(hDC, x + 1, y + 1, ETO_CLIPPED, lpRect, (LPCSTR)lpString, nLen, NULL);
		}
		else
		{
			TextOut(hDC, x + 1, y + 1, (LPCSTR)lpString, nLen);
		}
	}

	// 指定色で出力する
	SetTextColor(hDC, color);
	if (lpRect != NULL)
	{
		ExtTextOut(hDC, x, y, ETO_CLIPPED, lpRect, (LPCSTR)lpString, nLen, NULL);
	}
	else
	{
		TextOut(hDC, x, y, (LPCSTR)lpString, nLen);
	}
	
	// 出力位置を更新する
	SIZE size;
	GetTextExtentPoint32(hDC, (LPCSTR)lpString, nLen, &size);
	lpPoint->x += size.cx;
}

//
// コイン画像を表示する
//
// パラメータ
//   lpPoint        出力先座標
//   hDC			デバイスコンテキストのハンドル
//
void ShowCoinImage(LPPOINT lpPoint, HDC hDC)
{
#if _INMM_LOG_OUTPUT
#ifndef _INMM_PERF_LOG
	fprintf(fpLog, "  ShowCoinImage (%d,%d)\n", lpPoint->x, lpPoint->y);
#endif
#endif

	COLORREF crTransparent = GetPixel(hDCCoin, 0, 0);
	TransparentBlt(hDC, lpPoint->x + nCoinAdjustX, lpPoint->y + nCoinAdjustY, nCoinWidth, nCoinHeight, hDCCoin, 0, 0, nCoinWidth, nCoinHeight, (UINT)crTransparent);
	lpPoint->x += nCoinWidth;
}

//
// 改行位置を計算する (禁則処理対応版)
//
// パラメータ
//   lpBuffer		文字列処理バッファ
//   lpString		対象文字列
// 戻り値
//   処理したバイト数
//
int WINAPI CalcLineBreak(LPBYTE lpBuffer, LPCBYTE lpString)
{
#if _INMM_LOG_OUTPUT
#ifdef _INMM_PERF_LOG
	LARGE_INTEGER nFreq;
	LARGE_INTEGER nBefore;
	LARGE_INTEGER nAfter;
	memset(&nFreq, 0, sizeof(nFreq));
	memset(&nBefore, 0, sizeof(nBefore));
	memset(&nAfter, 0, sizeof(nAfter));
	QueryPerformanceFrequency(&nFreq);
	QueryPerformanceCounter(&nBefore);
#endif
#endif
	LPBYTE p = lpBuffer;
	LPCBYTE s = lpString;
	int len;

	while (*s != '\0')
	{
		// 行末禁則文字/制御文字の処理
		switch (*s)
		{
		case '(':
		case '<':
		case '[':
		case '`':
		case '{':
		case 0xA2: // ｢
			// 行末禁則文字ならば次の文字も一緒に処理する
			*p++ = *s++;
			continue;

		case 0x81:
			*p++ = *s++;
			switch (*s)
			{
			case 0x4D: // ｀
			case 0x65: // ‘
			case 0x67: // “
			case 0x69: // （
			case 0x6B: // 〔
			case 0x6D: // ［
			case 0x6F: // ｛
			case 0x71: // 〈
			case 0x73: // 《
			case 0x75: // 「
			case 0x77: // 『
			case 0x79: // 【
			case 0x83: // ＜
			case 0x8F: // ￥
			case 0x90: // ＄
				// 行末禁則文字ならば次の文字も一緒に処理する
				*p++ = *s++;
				continue;

			case 0x7B: // ＋
			case 0x7C: // －
			case 0x7D: // ±
				*p++ = *s++;
				// 連数字の分離禁則処理
				len = CalcNumberWordWrap(p, s);
				p += len;
				s += len;
				break;

			case 0x98: // §
				*p++ = *s++;
				// §の後に英大文字が続けばプリセットの色指定
				if ((*s >= 'A') && (*s <= 'Z'))
				{
					// §Wは白色なので分離禁則対象外
					if (*s == 'W')
					{
						*p++ = *s++;
						// 色指定の次の文字も一緒に処理する
						continue;
					}
					// 色指定後の分離禁則処理
					*p++ = *s++;
					len = CalcColorWordWrap(p, s);
					p += len;
					s += len;
					break;
				}
				// §§ならば元の色に戻す
				if ((*s == 0x81) && (*(s + 1) == 0x98))
				{
					*p++ = *s++;
					*p++ = *s++;
					// 色指定の次の文字も一緒に処理する
					continue;
				}
				// 次の3文字が0～7ならば直値の色指定
				if ((*s >= '0') && (*s <= '7') && (*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7'))
				{
					*p++ = *s++;
					*p++ = *s++;
					*p++ = *s++;
					// 色指定の次の文字も一緒に処理する
					continue;
				}
				break;

			default:
				// Shift_JISの2バイト目
				if (((*s >= 0x40) && (*s <= 0x7E)) || ((*s >= 0x80) && (*s <= 0xFC)))
				{
					*p++ = *s++;
					break;
				}
				// 不明な文字は1バイト文字として処理する
				break;
			}
			break;

		case 0x82:
			*p++ = *s++;
			// 2バイト数字
			if ((*s >= 0x16) && (*s <= 0x25))
			{
				*p++ = *s++;
				// 連数字の分離禁則処理
				len = CalcNumberWordWrap(p, s);
				p += len;
				s += len;
				break;
			}
			// Shift_JISの2バイト目
			if (((*s >= 0x40) && (*s <= 0x7E)) || ((*s >= 0x80) && (*s <= 0xFC)))
			{
				*p++ = *s++;
				break;
			}
			// 不明な文字は1バイト文字として処理する
			break;

		case 0x87:
			*p++ = *s++;
			if (*s == 0x80) // 〝
			{
				// 行末禁則文字ならば次の文字も一緒に処理する
				*p++ = *s++;
				continue;
			}
			// Shift_JISの2バイト目
			if (((*s >= 0x40) && (*s <= 0x7E)) || ((*s >= 0x80) && (*s <= 0xFC)))
			{
				*p++ = *s++;
				break;
			}
			// 不明な文字は1バイト文字として処理する
			break;

		case 0xA7: // ｧ
			*p++ = *s++;
			// ｧの後に英大文字が続けばプリセットの色指定
			if ((*s >= 'A') && (*s <= 'Z'))
			{
				// ｧWは白色なので分離禁則対象外
				if (*s == 'W')
				{
					*p++ = *s++;
					// 色指定の次の文字も一緒に処理する
					continue;
				}
				// 色指定後の分離禁則処理
				*p++ = *s++;
				len = CalcColorWordWrap(p, s);
				p += len;
				s += len;
				break;
			}
			// ｧｧならば元の色に戻す
			if (*s == 0xA7)
			{
				*p++ = *s++;
				// 色指定の次の文字も一緒に処理する
				continue;
			}
			// 次の3文字が0～7ならば直値の色指定
			if ((*s >= '0') && (*s <= '7') && (*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7'))
			{
				*p++ = *s++;
				*p++ = *s++;
				*p++ = *s++;
				// 色指定の次の文字も一緒に処理する
				continue;
			}
			break;

		case '%':
			*p++ = *s++;
			// 次の3文字が0～7ならば直値の色指定
			if ((*s >= '0') && (*s <= '7') && (*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7'))
			{
				*p++ = *s++;
				*p++ = *s++;
				*p++ = *s++;
				// 色指定の次の文字も一緒に処理する
				continue;
			}
			break;

		case '\\': // エスケープ記号
			*p++ = *s++;
			switch (*s)
			{
			case '$':
				// エスケープ記号の後に$が続けば単体の行末禁則文字として扱う
				*p++ = *s++;
				continue;

			case '%':
			case 0xA7: // ｧ
				// エスケープ記号の後に%/ｧが続けば単体の文字として扱う
				*p++ = *s++;
				break;

			case 0x81:
				if (*(s + 1) == 0x98)
				{
					// エスケープ記号の後に§が続けば単体の文字として扱う
					*p++ = *s++;
					*p++ = *s++;
				}
				else
				{
					// 単体の\は行末禁則文字として扱う
					continue;
				}
				break;

			default:
				// 単体の\は行末禁則文字として扱う
				continue;
			}
			break;

		case '0':
		case '1':
		case '2':
		case '3':
		case '4':
		case '5':
		case '6':
		case '7':
		case '8':
		case '9':
		case '-':
		case '+':
			*p++ = *s++;
			// 連数字の分離禁則処理
			len = CalcNumberWordWrap(p, s);
			p += len;
			s += len;
			break;

		case 0x0A:
		case 0x0D:
			// 禁則文字の次の改行コードは処理しない
			// これを入れないと色指定の直後の改行が処理されない
			break;

		case 0x83:
		case 0x84:
		case 0x85:
		case 0x86:
		case 0x88:
		case 0x89:
		case 0x8A:
		case 0x8B:
		case 0x8C:
		case 0x8D:
		case 0x8E:
		case 0x8F:
		case 0x90:
		case 0x91:
		case 0x92:
		case 0x93:
		case 0x94:
		case 0x95:
		case 0x96:
		case 0x97:
		case 0x98:
		case 0x99:
		case 0x9A:
		case 0x9B:
		case 0x9C:
		case 0x9D:
		case 0x9E:
		case 0x9F:
		case 0xE0:
		case 0xE1:
		case 0xE2:
		case 0xE3:
		case 0xE4:
		case 0xE5:
		case 0xE6:
		case 0xE7:
		case 0xE8:
		case 0xE9:
		case 0xEA:
		case 0xEB:
		case 0xEC:
		case 0xED:
		case 0xEE:
		case 0xEF:
		case 0xF0:
		case 0xF1:
		case 0xF2:
		case 0xF3:
		case 0xF4:
		case 0xF5:
		case 0xF6:
		case 0xF7:
		case 0xF8:
		case 0xF9:
		case 0xFA:
		case 0xFB:
		case 0xFC:
			*p++ = *s++;
			// Shift_JISの2バイト目
			if (((*s >= 0x40) && (*s <= 0x7E)) || ((*s >= 0x80) && (*s <= 0xFC)))
			{
				*p++ = *s++;
				break;
			}
			// 不明な文字は1バイト文字として処理する
			break;

		default:
			*p++ = *s++;
			break;
		}

		// 行頭禁則文字/分離禁則文字の処理
		switch (*s)
		{
		case '!':
		case '%':
		case '\'':
		case '-':
		case '.':
		case ',':
		case ')':
		case ':':
		case ';':
		case '=':
		case '>':
		case '?':
		case ']':
		case '}':
		case '~':
		case 0xA1: // ｡
		case 0xA3: // ｣
		case 0xA4: // ､
		case 0xA5: // ･
		case 0xA8: // ｨ
		case 0xA9: // ｩ
		case 0xAA: // ｪ
		case 0xAB: // ｫ
		case 0xAC: // ｬ
		case 0xAD: // ｭ
		case 0xAE: // ｮ
		case 0xAF: // ｯ
		case 0xB0: // ｰ
		case 0xDE: // ﾞ
		case 0xDF: // ﾟ
			// 次が行頭禁則文字ならば一緒に処理する
			continue;

		case '/':
			// 次が分離禁則文字ならばその次も含めて一緒に処理する
			*p++ = *s++;
			continue;

		case '\\':
			// 次が行頭禁則文字ならば一緒に処理する
			if (*(s + 1) == 0xA7) // ｧ
			{
				continue;
			}
			break;

		case 0x81:
			switch (*(s + 1))
			{
			case 0x41: // 、
			case 0x42: // 。
			case 0x43: // ，
			case 0x44: // ．
			case 0x45: // ・
			case 0x46: // ：
			case 0x47: // ；
			case 0x48: // ？
			case 0x49: // ！
			case 0x4A: // ゛
			case 0x4B: // ゜
			case 0x4C: // ´
			case 0x52: // ヽ
			case 0x53: // ヾ
			case 0x54: // ゝ
			case 0x55: // ゞ
			case 0x56: // 〃
			case 0x58: // 々
			case 0x5B: // ー
			case 0x5D: // ‐
			case 0x60: // ～
			case 0x66: // ’
			case 0x68: // ”
			case 0x6A: // ）
			case 0x6C: // 〕
			case 0x6E: // ］
			case 0x70: // ｝ 
			case 0x72: // 〉
			case 0x74: // 》
			case 0x76: // 」
			case 0x78: // 』
			case 0x7A: // 】
			case 0x81: // ＝
			case 0x84: // ＞
			case 0x8E: // ℃
			case 0x91: // ￠
			case 0x93: // ％
				// 次が行頭禁則文字ならば一緒に処理する
				continue;

			case 0x5C: // ―
			case 0x5E: // ／
			case 0x63: // …
			case 0x64: // ‥
				// 次が分離禁則文字ならばその次も含めて一緒に処理する
				*p++ = *s++;
				*p++ = *s++;
				continue;
			}
			break;

		case 0x82:
			switch (*(s + 1))
			{
			case 0x9F: // ぁ
			case 0xA1: // ぃ
			case 0xA3: // ぅ
			case 0xA5: // ぇ
			case 0xA7: // ぉ
			case 0xC1: // っ
			case 0xE1: // ゃ
			case 0xE3: // ゅ
			case 0xE5: // ょ
			case 0xEC: // ゎ
				// 次が行頭禁則文字ならば一緒に処理する
				continue;
			}
			break;

		case 0x83:
			switch (*(s + 1))
			{
			case 0x40: // ァ
			case 0x42: // ィ
			case 0x44: // ゥ
			case 0x46: // ェ
			case 0x48: // ォ
			case 0x62: // ッ
			case 0x83: // ャ
			case 0x85: // ュ
			case 0x87: // ョ
			case 0x8E: // ヮ
			case 0x95: // ヵ
			case 0x96: // ヶ
				// 次が行頭禁則文字ならば一緒に処理する
				continue;
			}
			break;

		case 0x87:
			if (*(s + 1) == 0x81) // 〟
			{
				// 次が行頭禁則文字ならば一緒に処理する
				continue;
			}
			break;
		}
		break;
	}
	*p = '\0';

#if _INMM_LOG_OUTPUT
#ifdef _INMM_PERF_LOG
	QueryPerformanceCounter(&nAfter);
	fprintf(fpLog, "[PERF] CalcLineBreak: %lf\n", (double)(nAfter.QuadPart - nBefore.QuadPart) * 1000 / nFreq.QuadPart);
#endif
	fprintf(fpLog, "CalcLineBreak: %s (%d) %s\n", lpBuffer, p - lpBuffer, lpString);
#endif

	return p - lpBuffer;
}

//
// 色指定後の分離禁則を計算する
//
// パラメータ
//   lpBuffer		文字列処理バッファ
//   lpString		対象文字列
// 戻り値
//   処理したバイト数
//
int CalcColorWordWrap(LPBYTE lpBuffer, LPCBYTE lpString)
{
	LPBYTE p = lpBuffer;
	LPCBYTE s = lpString;

	while (*s != '\0')
	{
		// 単語内の最大表示可能文字数を超える場合その場で区切る(保険)
		if (p - lpBuffer > nMaxWordChars)
		{
			break;
		}
		switch (*s)
		{
		case 0xA7: // ｧ
			// 色指定の直前で区切る
			if ((*(s + 1) >= 'A') && (*(s + 1) <= 'Z'))
			{
				break;
			}
			if (*(s + 1) == 0xA7)
			{
				break;
			}
			// 色指定でなければ通常の文字として扱う
			*p++ = *s++;
			continue;

		case 0x81:
			// §
			if (*(s + 1) == 0x98)
			{
				// 色指定の直前で区切る
				if ((*(s + 2) >= 'A') && (*(s + 2) <= 'Z'))
				{
					break;
				}
				if ((*(s + 2) == 0x81) && (*(s + 3) == 0x98))
				{
					break;
				}
				// 色指定でなければ通常の文字として扱う
				*p++ = *s++;
				*p++ = *s++;
				continue;
			}
			*p++ = *s++;
			// SHIFT_JISの2バイト目
			if (((*s >= 0x40) && (*s <= 0x7E)) || ((*s >= 0x80) && (*s <= 0xFC)))
			{
				*p++ = *s++;
				continue;
			}
			// 不明な文字は1バイト文字として処理する
			continue;

		case 0x82:
		case 0x83:
		case 0x84:
		case 0x85:
		case 0x86:
		case 0x87:
		case 0x88:
		case 0x89:
		case 0x8A:
		case 0x8B:
		case 0x8C:
		case 0x8D:
		case 0x8E:
		case 0x8F:
		case 0x90:
		case 0x91:
		case 0x92:
		case 0x93:
		case 0x94:
		case 0x95:
		case 0x96:
		case 0x97:
		case 0x98:
		case 0x99:
		case 0x9A:
		case 0x9B:
		case 0x9C:
		case 0x9D:
		case 0x9E:
		case 0x9F:
		case 0xE0:
		case 0xE1:
		case 0xE2:
		case 0xE3:
		case 0xE4:
		case 0xE5:
		case 0xE6:
		case 0xE7:
		case 0xE8:
		case 0xE9:
		case 0xEA:
		case 0xEB:
		case 0xEC:
		case 0xED:
		case 0xEE:
		case 0xEF:
		case 0xF0:
		case 0xF1:
		case 0xF2:
		case 0xF3:
		case 0xF4:
		case 0xF5:
		case 0xF6:
		case 0xF7:
		case 0xF8:
		case 0xF9:
		case 0xFA:
		case 0xFB:
		case 0xFC:
			*p++ = *s++;
			// SHIFT_JISの2バイト目
			if (((*s >= 0x40) && (*s <= 0x7E)) || ((*s >= 0x80) && (*s <= 0xFC)))
			{
				*p++ = *s++;
				continue;
			}
			// 不明な文字は1バイト文字として処理する
			continue;

		case 0x01:
		case 0x02:
		case 0x03:
		case 0x04:
		case 0x05:
		case 0x06:
		case 0x07:
		case 0x08:
		case 0x09:
		case 0x0A:
		case 0x0B:
		case 0x0C:
		case 0x0D:
		case 0x0E:
		case 0x0F:
		case 0x10:
		case 0x11:
		case 0x12:
		case 0x13:
		case 0x14:
		case 0x15:
		case 0x16:
		case 0x17:
		case 0x18:
		case 0x19:
		case 0x1A:
		case 0x1B:
		case 0x1C:
		case 0x1D:
		case 0x1E:
		case 0x1F:
			// 制御文字があれば直前で区切る
			break;

		default:
			*p++ = *s++;
			continue;
		}
		break;
	}

#if _INMM_LOG_OUTPUT
#ifndef _INMM_PERF_LOG
	*p = '\0';
	fprintf(fpLog, "CalcColorWordWrap: %s (%d) %s\n", lpBuffer, p - lpBuffer, lpString);
#endif
#endif

	return p - lpBuffer;
}

//
// 連数字の分離禁則を計算する
//
// パラメータ
//   lpBuffer		文字列処理バッファ
//   lpString		対象文字列
// 戻り値
//   処理したバイト数
//
int CalcNumberWordWrap(LPBYTE lpBuffer, LPCBYTE lpString)
{
	LPBYTE p = lpBuffer;
	LPCBYTE s = lpString;

	while (*s != '\0')
	{
		// 単語内の最大表示可能文字数を超える場合その場で区切る(保険)
		if (p - lpBuffer > nMaxWordChars)
		{
			break;
		}
		switch (*s)
		{
		case '0':
		case '1':
		case '2':
		case '3':
		case '4':
		case '5':
		case '6':
		case '7':
		case '8':
		case '9':
		case ',':
		case '.':
			// 数字または数字を構成する記号ならば続けて処理する
			*p++ = *s++;
			continue;

		case '%':
			if ((*(s + 1) < '0') || (*(s + 1) > '9') || (*(s + 2) < '0') || (*(s + 2) > '9') || (*(s + 3) < '0') || (*(s + 3) > '9'))
			{
				// 数字の後ろの%はまとめて処理する
				*p++ = *s++;
			}
			break;

		case '\\':
			if (*(s + 1) == '%')
			{
				// 数字の後ろの%はまとめて処理する
				*p++ = *s++;
				*p++ = *s++;
			}
			break;

		case 0x81:
			switch (*(s + 1))
			{
			case 0x43: // ，
			case 0x44: // ．
				*p++ = *s++;
				*p++ = *s++;
				continue;

			case 0x93: // ％
				// 数字の後ろの%はまとめて処理する
				*p++ = *s++;
				*p++ = *s++;
				break;
			}
			break;

		case 0x82:
			// 2バイト数字
			if ((*(s + 1) >= 0x16) && (*(s + 1) <= 0x25))
			{
				*p++ = *s++;
				*p++ = *s++;
				continue;
			}
			break;
		}
		break;
	}

#if _INMM_LOG_OUTPUT
#ifndef _INMM_PERF_LOG
	*p = '\0';
	fprintf(fpLog, "CalcNumberWordWrap: %s (%d) %s\n", lpBuffer, p - lpBuffer, lpString);
#endif
#endif

	return p - lpBuffer;
}

//
// 文字列長を取得する(最大指定あり)
//
// パラメータ
//   lpString		対象文字列
//   nMax			最大文字列長
// 戻り値
//   文字列長
//
int WINAPI strnlen0(LPCBYTE lpString, int nMax)
{
	if (lpString == NULL)
	{
		return 0;
	}

	LPCBYTE s = lpString;
	int len = 0;
	while (*s != '\0')
	{
		// Shift_JISの1バイト目
		if (((*s >= 0x81) && (*s <= 0x9F)) || ((*s >= 0xE0) && (*s <= 0xFC)))
		{
			// Shift_JISの2バイト目
			if (((*(s + 1) >= 0x40) && (*(s + 1) <= 0x7E)) || ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xFC)))
			{
				if (len + 1 >= nMax)
				{
					break;
				}
				s += 2;
				len += 2;
				continue;
			}
		}
		// それ以外の文字は1バイト文字として処理する
		if (len >= nMax)
		{
			break;
		}
		s++;
		len++;
	}

#if _INMM_LOG_OUTPUT
	fprintf(fpLog, "strnlen0: %s (%d) => %d\n", lpString, nMax, len);
#endif

	return len;
}

// winmm.dllのmciSendCommand APIへ転送
MCIERROR WINAPI _mciSendCommandA(MCIDEVICEID IDDevice, UINT uMsg, DWORD fdwCommand, DWORD dwParam)
{
	return lpfnMciSendCommandA(IDDevice, uMsg, fdwCommand, dwParam);
}

// winmm.dllのtimeBeginPeriod APIへ転送
MMRESULT WINAPI _timeBeginPeriod(UINT uPeriod)
{
	return lpfnTimeBeginPeriod(uPeriod);
}

// winmm.dllのtimeGetDevCaps APIへ転送
MMRESULT WINAPI _timeGetDevCaps(LPTIMECAPS ptc, UINT cbtc)
{
	return lpfnTimeGetDevCaps(ptc, cbtc);
}

// winmm.dllのtimeGetTime APIへ転送
DWORD WINAPI _timeGetTime(VOID)
{
	return lpfnTimeGetTime();
}

// winmm.dllのtimeKillEvent APIへ転送
MMRESULT WINAPI _timeKillEvent(UINT uTimerID)
{
	return lpfnTimeKillEvent(uTimerID);
}

// winmm.dllのtimeSetEvent APIへ転送
MMRESULT WINAPI _timeSetEvent(UINT uDelay, UINT uResolution, LPTIMECALLBACK lpTimeProc, DWORD dwUser, UINT fuEvent)
{
	return lpfnTimeSetEvent(uDelay, uResolution, lpTimeProc, dwUser, fuEvent);
}

//
// 初期化処理
//
void Init()
{
	// winmm.dllのAPI転送関数を設定する
	hWinMmDll = LoadLibrary("winmm.dll");
	lpfnMciSendCommandA = (LPFNMCISENDCOMMANDA)GetProcAddress(hWinMmDll, "mciSendCommandA");
	lpfnTimeBeginPeriod = (LPFNTIMEBEGINPERIOD)GetProcAddress(hWinMmDll, "timeBeginPeriod");
	lpfnTimeGetDevCaps = (LPFNTIMEGETDEVCAPS)GetProcAddress(hWinMmDll, "timeGetDevCaps");
	lpfnTimeGetTime = (LPFNTIMEGETTIME)GetProcAddress(hWinMmDll, "timeGetTime");
	lpfnTimeKillEvent = (LPFNTIMEKILLEVENT)GetProcAddress(hWinMmDll, "timeKillEvent");
	lpfnTimeSetEvent = (LPFNTIMESETEVENT)GetProcAddress(hWinMmDll, "timeSetEvent");

#if _INMM_LOG_OUTPUT
	fopen_s(&fpLog, "_inmm.log", "w");
#endif

	hDesktopDC = GetWindowDC(HWND_DESKTOP);

	LoadIniFile();
	InitFont();
	InitPalette();
	InitCoinImage();
}

//
// iniファイルから設定を読み込む
//
void LoadIniFile()
{
	char szIniFileName[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, szIniFileName);
	lstrcat(szIniFileName, "\\_inmm.ini");

	FILE *fp;
	fopen_s(&fp, szIniFileName, "r");
	if (fp == NULL)
	{
		return;
	}

	const int nLineBufSize = 512;
	char szLineBuf[nLineBufSize];
	int nMagicCode = 0;

	while (!feof(fp))
	{
		if (!ReadLine(fp, szLineBuf, nLineBufSize))
		{
			break;
		}
		char *p = szLineBuf;
		// 空行/コメント行/セクション行は読み飛ばす
		if ((*p == '\0') || (*p == '#') || (*p == '['))
		{
			continue;
		}
		char *s = p;
		while ((*s != '=') && (*s != '\0'))
		{
			s++;
		}
		// 定義行の形式でない場合は無視する
		if (*s == '\0')
		{
			continue;
		}
		*s++ = '\0';
		// 色指定
		if (!strncmp(p, "Color", 5) && (*(p + 5) >= 'A') && (*(p + 5) <= 'Z') && (*s == '$'))
		{
			int nColorIndex = *(p + 5) - 'A';
			DWORD color = strtoul(s + 1, NULL, 16);
			colorTable[nColorIndex].r = (BYTE)(color & 0x000000FF);
			colorTable[nColorIndex].g = (BYTE)((color & 0x0000FF00) >> 8);
			colorTable[nColorIndex].b = (BYTE)((color & 0x00FF0000) >> 16);
		}
		else if (!lstrcmp(p, "MagicCode"))
		{
			nMagicCode = strtol(s, NULL, 0);
			if ((nMagicCode < 0) || (nMagicCode >= nMaxFontTable))
			{
				nMagicCode = 0;
			}
		}
		else if (!lstrcmp(p, "Font"))
		{
			lstrcpy(fontTable[nMagicCode].szFaceName, s);
		}
		else if (!lstrcmp(p, "Height"))
		{
			fontTable[nMagicCode].nHeight = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "Bold"))
		{
			fontTable[nMagicCode].nBold = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "AdjustX"))
		{
			fontTable[nMagicCode].nAdjustX = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "AdjustY"))
		{
			fontTable[nMagicCode].nAdjustY = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "Coin"))
		{
			lstrcpy(szCoinFileName, s);
		}
		else if (!lstrcmp(p, "CoinAdjustX"))
		{
			nCoinAdjustX = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "CoinAdjustY"))
		{
			nCoinAdjustY = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "MaxLines"))
		{
			nMaxLines = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "MaxWordChars"))
		{
			nMaxWordChars = strtol(s, NULL, 0);
		}
	}
	fclose(fp);
}

//
// フォントの初期化
//
void InitFont()
{
	LOGFONT lf;
	memset(&lf, 0, sizeof(lf));
	lf.lfWidth = 0;
	lf.lfEscapement = 0;
	lf.lfOrientation = 0;
	lf.lfItalic = FALSE;
	lf.lfUnderline = FALSE;
	lf.lfStrikeOut = FALSE;
	lf.lfCharSet = DEFAULT_CHARSET;
	lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
	lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
	lf.lfQuality = DEFAULT_QUALITY;
	lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;

	for (int i = 0; i < nMaxFontTable; i++)
	{
		if (fontTable[i].szFaceName[0] == '\0')
		{
			if (i == 0)
			{
				lstrcpy(fontTable[i].szFaceName, "MS UI Gothic");
			}
			else
			{
				continue;
			}
		}
		lf.lfHeight = fontTable[i].nHeight;
		lf.lfWeight = (fontTable[i].nBold != 0) ? FW_BOLD : FW_NORMAL;
		lstrcpy(lf.lfFaceName, fontTable[i].szFaceName);
		fontTable[i].hFont = CreateFontIndirect(&lf);
	}
}

//
// パレットの初期化
//
void InitPalette()
{
	LOGPALETTE *lpPalette = (LOGPALETTE *)malloc(sizeof(LOGPALETTE) + sizeof(PALETTEENTRY) * (nMaxColorTable - 1));
	lpPalette->palVersion = 0x0300;
	lpPalette->palNumEntries = nMaxColorTable;
	for (int i = 0; i < nMaxColorTable; i++)
	{
		lpPalette->palPalEntry[i].peRed = colorTable[i].r;
		lpPalette->palPalEntry[i].peGreen = colorTable[i].g;
		lpPalette->palPalEntry[i].peBlue = colorTable[i].b;
		lpPalette->palPalEntry[i].peFlags = 0;
	}
	hPalette = CreatePalette(lpPalette);
	free(lpPalette);
	SelectPalette(hDesktopDC, hPalette, FALSE);
	RealizePalette(hDesktopDC);
}

//
// コイン画像の初期化
//
void InitCoinImage()
{
	hBitmapCoin = (HBITMAP)LoadImage(NULL, szCoinFileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
	if (hBitmapCoin == NULL)
	{
		return;
	}
	hDCCoin = CreateCompatibleDC(hDesktopDC);
	hBitmapCoinPrev = (HBITMAP)SelectObject(hDCCoin, hBitmapCoin);
}

//
// 終了処理
//
void Terminate()
{
#if _INMM_LOG_OUTPUT
	fclose(fpLog);
#endif
	SelectObject(hDCCoin, hBitmapCoinPrev);
	DeleteObject(hBitmapCoin);
	DeleteDC(hDCCoin);
	DeleteObject((HGDIOBJ)hPalette);
	for (int i = 0; i < nMaxFontTable; i++)
	{
		if (fontTable[i].hFont != NULL)
		{
			DeleteObject((HGDIOBJ)fontTable[i].hFont);
		}
	}
	ReleaseDC(HWND_DESKTOP, hDesktopDC);
	FreeLibrary(hWinMmDll);
}

//
// ファイルから1行を読み込み、末尾の改行文字を削除する
//
// パラメータ
//   fp				ファイルポインタ
//   buf			文字列バッファ
//   len			文字列バッファの最大長
//
bool ReadLine(FILE *fp, char *buf, int len)
{
	if (fgets(buf, len, fp) == NULL)
	{
		return false;
	}
	char *p = strchr(buf, '\n');
	if (*p != NULL)
	{
		*p = '\0';
	}
	return true;
}
