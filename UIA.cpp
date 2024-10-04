#include <Windows.h>
#include <UIAutomation.h>
#include <iostream>
#include <string>
#include <oleauto.h>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "UIAutomationCore.lib")

std::wstring GetElementName(IUIAutomationElement* pElement) {
    BSTR bstrName;
    pElement->get_CurrentName(&bstrName);
    std::wstring name(bstrName, SysStringLen(bstrName));
    SysFreeString(bstrName);
    return name;
}

std::wstring GetElementClassName(IUIAutomationElement* pElement) {
    BSTR bstrClassName;
    pElement->get_CurrentClassName(&bstrClassName);
    std::wstring className(bstrClassName, SysStringLen(bstrClassName));
    SysFreeString(bstrClassName);
    return className;
}

int GetElementControlType(IUIAutomationElement* pElement) {
    int controlType;
    pElement->get_CurrentControlType(&controlType);
    return controlType;
}

void PrintElementInfo(IUIAutomationElement* pElement, const std::wstring& prefix) {
    std::wcout << prefix << L"Name: " << GetElementName(pElement) << std::endl;
    std::wcout << prefix << L"ClassName: " << GetElementClassName(pElement) << std::endl;
    std::wcout << prefix << L"ControlType: " << GetElementControlType(pElement) << std::endl;
}

IUIAutomationElement* FindGridElement(IUIAutomationElement* pParent, IUIAutomation* pAutomation) {
    IUIAutomationCondition* pCondition = NULL;
    IUIAutomationElement* pGrid = NULL;

    // Try to find DataGrid control
    VARIANT varProp;
    VariantInit(&varProp);
    varProp.vt = VT_BSTR;
    varProp.bstrVal = SysAllocString(L"DataGridView");

    HRESULT hr = pAutomation->CreatePropertyCondition(UIA_ClassNamePropertyId, varProp, &pCondition);
    if (SUCCEEDED(hr)) {
        pParent->FindFirst(TreeScope_Descendants, pCondition, &pGrid);
        pCondition->Release();
    }

    VariantClear(&varProp);

    // If DataGrid not found, try to find any grid-like control
    if (!pGrid) {
        varProp.vt = VT_I4;
        varProp.lVal = UIA_DataGridControlTypeId;
        hr = pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, varProp, &pCondition);
        if (SUCCEEDED(hr)) {
            pParent->FindFirst(TreeScope_Descendants, pCondition, &pGrid);
            pCondition->Release();
        }
    }

    return pGrid;
}

int main()
{
    CoInitialize(NULL);
    
    IUIAutomation* pAutomation = NULL;
    HRESULT hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, 
                                  __uuidof(IUIAutomation), (void**)&pAutomation);

    if (SUCCEEDED(hr))
    {
        IUIAutomationElement* pRoot = NULL;
        hr = pAutomation->GetRootElement(&pRoot);

        if (SUCCEEDED(hr) && pRoot)
        {
            IUIAutomationElement* pWindow = NULL;
            IUIAutomationCondition* pWindowCondition = NULL;

            VARIANT varWindowProp;
            VariantInit(&varWindowProp);
            varWindowProp.vt = VT_BSTR;
            varWindowProp.bstrVal = SysAllocString(L"Grid Control Demo");

            hr = pAutomation->CreatePropertyCondition(UIA_NamePropertyId, varWindowProp, &pWindowCondition);
            
            if (SUCCEEDED(hr))
            {
                pRoot->FindFirst(TreeScope_Children, pWindowCondition, &pWindow);
                pWindowCondition->Release();
                VariantClear(&varWindowProp);

                if (pWindow)
                {
                    std::wcout << L"Found window:\n";
                    PrintElementInfo(pWindow, L"  ");

                    IUIAutomationElement* pGrid = FindGridElement(pWindow, pAutomation);

                    if (pGrid)
                    {
                        std::wcout << L"Found grid element:\n";
                        PrintElementInfo(pGrid, L"  ");

                        pGrid->Release();
                    }
                    else
                    {
                        std::wcout << L"Grid element not found\n";
                    }

                    pWindow->Release();
                }
                else
                {
                    std::wcout << L"Window 'Grid Control Demo' not found\n";
                }
            }

            pRoot->Release();
        }
        
        pAutomation->Release();
    }
    
    CoUninitialize();
    return 0;
}