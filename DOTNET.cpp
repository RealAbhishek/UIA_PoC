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
                    IUIAutomationElementArray* pChildren = NULL;
                    pWindow->FindAll(TreeScope_Children, NULL, &pChildren);

                    if (pChildren)
                    {
                        int childCount;
                        pChildren->get_Length(&childCount);
                        std::wcout << L"Window has " << childCount << L" child elements.\n";

                        for (int i = 0; i < childCount; i++)
                        {
                            IUIAutomationElement* pChild = NULL;
                            pChildren->GetElement(i, &pChild);
                            if (pChild)
                            {
                                std::wcout << L"Child " << i << L":\n";
                                PrintElementInfo(pChild, L"  ");
                                pChild->Release();
                            }
                        }

                        pChildren->Release();
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