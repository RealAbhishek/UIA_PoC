#define UNICODE
#define _UNICODE
#define _WIN32_WINNT 0x0600 // Targeting Windows Vista or later
#define _WIN32_IE 0x0600    // Targeting IE 6.0 or later

#include <windows.h>
#include <commctrl.h>
#include <wchar.h>          // For wide-character functions
#include <strsafe.h>        // For StringCchCopyW
#include <shellapi.h>       // For CommandLineToArgvW
#include <string>
#include <vector>
#include <memory>
#include <sstream>
#include <algorithm>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shell32.lib") // Link against Shell32.lib

// Forward declarations
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK ListViewProc(HWND, UINT, WPARAM, LPARAM);

// Global variables
WNDPROC g_OldListViewProc = NULL;
int g_TotalColumnWidth = 0;

// Structure to hold table dimensions
struct TableDimensions {
    int numColumns;
    int numRows;
};

// Function to parse command-line arguments
std::unique_ptr<TableDimensions> ParseCommandLineArgs(LPWSTR lpCmdLine) {
    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLine(), &argc);
    if (!argv) {
        // Allocation failed
        return nullptr;
    }

    std::unique_ptr<TableDimensions> dims = std::make_unique<TableDimensions>();
    // Set default values
    dims->numColumns = 3;
    dims->numRows = 50000;

    if (argc >= 3) { // argv[0] is the program name
        try {
            int cols = std::stoi(argv[1]);
            int rows = std::stoi(argv[2]);

            // Validate the arguments
            if (cols > 0 && rows > 0) {
                dims->numColumns = cols;
                dims->numRows = rows;
            } else {
                MessageBox(NULL, L"Invalid arguments. Using default values (3 columns, 50000 rows).", L"Invalid Arguments", MB_ICONWARNING | MB_OK);
            }
        } catch (...) {
            MessageBox(NULL, L"Failed to parse arguments. Using default values (3 columns, 50000 rows).", L"Parsing Error", MB_ICONWARNING | MB_OK);
        }
    } else if (argc > 1) {
        MessageBox(NULL, L"Insufficient arguments. Using default values (3 columns, 50000 rows).", L"Insufficient Arguments", MB_ICONWARNING | MB_OK);
    }

    LocalFree(argv);
    return dims;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow)
{
    // Parse command-line arguments
    std::unique_ptr<TableDimensions> tableDims = ParseCommandLineArgs(lpCmdLine);
    if (!tableDims) {
        MessageBox(NULL, L"Failed to parse command-line arguments.", L"Error", MB_ICONERROR | MB_OK);
        return 0;
    }

    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_LISTVIEW_CLASSES;
    InitCommonControlsEx(&icex);

    WNDCLASSEX wc = {0};
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.style = 0;
    wc.lpfnWndProc = WndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = sizeof(TableDimensions*); // Allocate extra bytes to store pointer
    wc.hInstance = hInstance;
    wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
    wc.lpszMenuName = NULL;
    wc.lpszClassName = L"TableAppClass";
    wc.hIconSm = LoadIcon(NULL, IDI_APPLICATION);

    if(!RegisterClassEx(&wc))
    {
        MessageBox(NULL, L"Window Registration Failed!", L"Error!", MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    HWND hwnd = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        L"TableAppClass",
        L"Table App",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 1024, 768, // Increased initial size
        NULL, NULL, hInstance, tableDims.get());

    if(hwnd == NULL)
    {
        MessageBox(NULL, L"Window Creation Failed!", L"Error!", MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    // Transfer ownership of tableDims to the window data
    SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)tableDims.release());

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG Msg;
    while(GetMessage(&Msg, NULL, 0, 0) > 0)
    {
        TranslateMessage(&Msg);
        DispatchMessage(&Msg);
    }
    return (int)Msg.wParam;
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    static HWND hListView = NULL;
    static TableDimensions* pTableDims = nullptr;

    switch(msg)
    {
    case WM_CREATE:
        {
            LPCREATESTRUCT pcs = (LPCREATESTRUCT)lParam;
            pTableDims = (TableDimensions*)pcs->lpCreateParams;

            hListView = CreateWindowEx(0, WC_LISTVIEW, L"", 
                WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_OWNERDATA | WS_HSCROLL, 
                10, 10, 780, 580, 
                hwnd, (HMENU)1, GetModuleHandle(NULL), NULL);

            ListView_SetExtendedListViewStyle(hListView, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_HEADERDRAGDROP);

            // Subclass the ListView to handle custom drawing
            g_OldListViewProc = (WNDPROC)SetWindowLongPtr(hListView, GWLP_WNDPROC, (LONG_PTR)ListViewProc);

            // Set up virtual columns
            LVCOLUMN lvc = {0};
            lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
            lvc.fmt = LVCFMT_LEFT;
            lvc.cx = 100;
            lvc.pszText = LPSTR_TEXTCALLBACK;

            for (int i = 0; i < pTableDims->numColumns; i++)
            {
                lvc.iSubItem = i;
                ListView_InsertColumn(hListView, i, &lvc);
                g_TotalColumnWidth += lvc.cx;
            }

            // Set the number of items
            ListView_SetItemCountEx(hListView, pTableDims->numRows, LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL);
        }
        break;
    case WM_SIZE:
        {
            if (hListView)
            {
                RECT rcClient;
                GetClientRect(hwnd, &rcClient);
                int width = rcClient.right - rcClient.left - 20;
                int height = rcClient.bottom - rcClient.top - 20;
                SetWindowPos(hListView, NULL, 10, 10, width, height, SWP_NOZORDER);
            }
        }
        break;
    case WM_NOTIFY:
        {
            LPNMHDR pnmhdr = reinterpret_cast<LPNMHDR>(lParam);
            if (pnmhdr->idFrom == 1)
            {
                switch (pnmhdr->code)
                {
                case LVN_GETDISPINFO:
                    {
                        NMLVDISPINFO* pDispInfo = reinterpret_cast<NMLVDISPINFO*>(lParam);
                        if(pDispInfo->item.mask & LVIF_TEXT)
                        {
                            int itemIndex = pDispInfo->item.iItem;
                            int subItem = pDispInfo->item.iSubItem;

                            std::wstring text = L"Cell " + std::to_wstring(itemIndex + 1) + L"," + std::to_wstring(subItem + 1);

                            StringCchCopyW(pDispInfo->item.pszText, pDispInfo->item.cchTextMax, text.c_str());
                        }
                    }
                    break;
                }
            }
        }
        break;
    case WM_CLOSE:
        DestroyWindow(hwnd);
        break;
    case WM_DESTROY:
        {
            TableDimensions* pDims = (TableDimensions*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
            if(pDims)
            {
                delete pDims;
            }
            PostQuitMessage(0);
        }
        break;
    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

LRESULT CALLBACK ListViewProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_NOTIFY:
        {
            LPNMHDR pnmhdr = reinterpret_cast<LPNMHDR>(lParam);
            if (pnmhdr->code == HDN_ITEMCHANGED)
            {
                LPNMHEADER phdr = reinterpret_cast<LPNMHEADER>(lParam);
                if (phdr->pitem->mask & HDI_WIDTH)
                {
                    // Recalculate total column width
                    g_TotalColumnWidth = 0;
                    int columnCount = Header_GetItemCount(pnmhdr->hwndFrom);
                    for (int i = 0; i < columnCount; i++)
                    {
                        HDITEM hdi = {0};
                        hdi.mask = HDI_WIDTH;
                        Header_GetItem(pnmhdr->hwndFrom, i, &hdi);
                        g_TotalColumnWidth += hdi.cxy;
                    }
                }
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);

            // Custom drawing for header
            HWND hHeader = ListView_GetHeader(hwnd);
            if (hHeader)
            {
                RECT headerRect;
                GetClientRect(hHeader, &headerRect);
                
                // Draw the header background
                FillRect(hdc, &headerRect, (HBRUSH)(COLOR_BTNFACE + 1));

                // Draw each header item
                int columnCount = Header_GetItemCount(hHeader);
                int xPos = 0;
                for (int i = 0; i < columnCount; i++)
                {
                    RECT itemRect;
                    Header_GetItemRect(hHeader, i, &itemRect);

                    // Adjust for horizontal scrolling
                    itemRect.left += xPos;
                    itemRect.right += xPos;

                    // Draw the column text
                    std::wstring columnText = L"Column " + std::to_wstring(i + 1);
                    DrawText(hdc, columnText.c_str(), -1, &itemRect, DT_SINGLELINE | DT_VCENTER | DT_CENTER);

                    // Draw a line to separate columns
                    MoveToEx(hdc, itemRect.right, itemRect.top, NULL);
                    LineTo(hdc, itemRect.right, itemRect.bottom);

                    xPos += itemRect.right - itemRect.left;
                }
            }

            EndPaint(hwnd, &ps);
        }
        return 0;
    }

    return CallWindowProc(g_OldListViewProc, hwnd, msg, wParam, lParam);
}