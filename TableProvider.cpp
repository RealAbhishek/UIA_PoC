#define UNICODE
#define _UNICODE
#include <windows.h>
#include <commctrl.h>
#include <string>
#include <UIAutomationCore.h>
#include <UIAutomation.h>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "UIAutomationCore.lib")


class ListViewProvider : 
    public IRawElementProviderSimple,
    public ITableProvider,
    public ISelectionProvider,
    public IScrollProvider
{
private:
    HWND m_hwnd;
    HWND m_hListView;
    long m_refCount;

public:
    ListViewProvider(HWND hwnd, HWND hListView) : m_hwnd(hwnd), m_hListView(hListView), m_refCount(1) {}

    // IUnknown methods
    ULONG STDMETHODCALLTYPE AddRef() {
        return InterlockedIncrement(&m_refCount);
    }

    ULONG STDMETHODCALLTYPE Release() {
        long val = InterlockedDecrement(&m_refCount);
        if (val == 0) {
            delete this;
        }
        return val;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface) {
        if (riid == __uuidof(IUnknown))
            *ppInterface = static_cast<IRawElementProviderSimple*>(this);
        else if (riid == __uuidof(IRawElementProviderSimple))
            *ppInterface = static_cast<IRawElementProviderSimple*>(this);
        else if (riid == __uuidof(ITableProvider))
            *ppInterface = static_cast<ITableProvider*>(this);
        else if (riid == __uuidof(ISelectionProvider))
            *ppInterface = static_cast<ISelectionProvider*>(this);
        else if (riid == __uuidof(IScrollProvider))
            *ppInterface = static_cast<IScrollProvider*>(this);
        else {
            *ppInterface = NULL;
            return E_NOINTERFACE;
        }
        AddRef();
        return S_OK;
    }

    // IRawElementProviderSimple methods
    HRESULT STDMETHODCALLTYPE get_ProviderOptions(ProviderOptions* pRetVal) {
        *pRetVal = ProviderOptions_ServerSideProvider;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetPatternProvider(PATTERNID patternId, IUnknown** pRetVal) {
        *pRetVal = NULL;
        if (patternId == UIA_TablePatternId || 
            patternId == UIA_SelectionPatternId || 
            patternId == UIA_ScrollPatternId)
        {
            return QueryInterface(IID_IUnknown, (void**)pRetVal);
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetPropertyValue(PROPERTYID propertyId, VARIANT* pRetVal) {
        VariantInit(pRetVal);
        switch (propertyId) {
            case UIA_ControlTypePropertyId:
                pRetVal->vt = VT_I4;
                pRetVal->lVal = UIA_ListControlTypeId;
                break;
            // Implement other properties...
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE get_HostRawElementProvider(IRawElementProviderSimple** pRetVal) {
        return UiaHostProviderFromHwnd(m_hwnd, pRetVal);
    }

    // ITableProvider methods
    HRESULT STDMETHODCALLTYPE GetRowHeaders(SAFEARRAY** pRetVal) {
        *pRetVal = NULL;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetColumnHeaders(SAFEARRAY** pRetVal) {
        *pRetVal = NULL;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE get_RowOrColumnMajor(RowOrColumnMajor* pRetVal) {
        *pRetVal = RowOrColumnMajor_RowMajor;
        return S_OK;
    }

    // ISelectionProvider methods
    HRESULT STDMETHODCALLTYPE GetSelection(SAFEARRAY** pRetVal) {
        // Implement selection retrieval
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE get_CanSelectMultiple(BOOL* pRetVal) {
        *pRetVal = TRUE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE get_IsSelectionRequired(BOOL* pRetVal) {
        *pRetVal = FALSE;
        return S_OK;
    }

    // IScrollProvider methods
    HRESULT STDMETHODCALLTYPE Scroll(ScrollAmount horizontalAmount, ScrollAmount verticalAmount) {
        // Implement scrolling
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE SetScrollPercent(double horizontalPercent, double verticalPercent) {
        // Implement setting scroll percent
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE get_HorizontalScrollPercent(double* pRetVal) {
        // Implement getting horizontal scroll percent
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE get_VerticalScrollPercent(double* pRetVal) {
        // Implement getting vertical scroll percent
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE get_HorizontalViewSize(double* pRetVal) {
        // Implement getting horizontal view size
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE get_VerticalViewSize(double* pRetVal) {
        // Implement getting vertical view size
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE get_HorizontallyScrollable(BOOL* pRetVal) {
        *pRetVal = TRUE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE get_VerticallyScrollable(BOOL* pRetVal) {
        *pRetVal = TRUE;
        return S_OK;
    }
};

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow)
{
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_LISTVIEW_CLASSES;
    InitCommonControlsEx(&icex);

    WNDCLASSEX wc = {0};
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.style = 0;
    wc.lpfnWndProc = WndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = 0;
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
        CW_USEDEFAULT, CW_USEDEFAULT, 500, 400,
        NULL, NULL, hInstance, NULL);

    if(hwnd == NULL)
    {
        MessageBox(NULL, L"Window Creation Failed!", L"Error!", MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG Msg;
    while(GetMessage(&Msg, NULL, 0, 0) > 0)
    {
        TranslateMessage(&Msg);
        DispatchMessage(&Msg);
    }
    return Msg.wParam;
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    static HWND hListView = NULL;
    static ListViewProvider* pProvider = NULL;

    switch(msg)
    {
    case WM_CREATE:
        {
            hListView = CreateWindowEx(0, WC_LISTVIEW, L"", 
                WS_CHILD | WS_VISIBLE | LVS_REPORT, 
                10, 10, 480, 350, 
                hwnd, (HMENU)1, GetModuleHandle(NULL), NULL);

            ListView_SetExtendedListViewStyle(hListView, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

            LVCOLUMN lvc;
            lvc.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;

            lvc.iSubItem = 0;
            lvc.pszText = const_cast<LPWSTR>(L"Column 1");
            lvc.cx = 100;
            ListView_InsertColumn(hListView, 0, &lvc);

            lvc.iSubItem = 1;
            lvc.pszText = const_cast<LPWSTR>(L"Column 2");
            lvc.cx = 100;
            ListView_InsertColumn(hListView, 1, &lvc);

            lvc.iSubItem = 2;
            lvc.pszText = const_cast<LPWSTR>(L"Column 3");
            lvc.cx = 100;
            ListView_InsertColumn(hListView, 2, &lvc);

            LVITEM lvi;
            lvi.mask = LVIF_TEXT;
            for(int i = 0; i < 50000; i++)
            {
                lvi.iItem = i;
                lvi.iSubItem = 0;
                std::wstring text = L"Item " + std::to_wstring(i+1);
                lvi.pszText = const_cast<LPWSTR>(text.c_str());
                ListView_InsertItem(hListView, &lvi);

                for(int j = 1; j < 3; j++)
                {
                    ListView_SetItemText(hListView, i, j, const_cast<LPWSTR>(L"SubItem"));
                }
            }

            pProvider = new ListViewProvider(hwnd, hListView);
        }
        break;
    case WM_GETOBJECT:
        if (lParam == UiaRootObjectId) {
            IRawElementProviderSimple* pElement = NULL;
            pProvider->QueryInterface(IID_PPV_ARGS(&pElement));
            return UiaReturnRawElementProvider(hwnd, wParam, lParam, pElement);
        }
        break;
    case WM_CLOSE:
        DestroyWindow(hwnd);
        break;
    case WM_DESTROY:
        if (pProvider) {
            pProvider->Release();
            pProvider = NULL;
        }
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}