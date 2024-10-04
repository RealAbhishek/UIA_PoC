/*
#include <Windows.h>
#include <UIAutomation.h>
#include <iostream>
#include <string>
#include <oleauto.h>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "UIAutomationCore.lib")

std::wstring GetCellValue(IUIAutomationElement* pCell) {
    BSTR bstrValue;
    pCell->get_CurrentName(&bstrValue);
    std::wstring value(bstrValue, SysStringLen(bstrValue));
	
	std::wcout << L"Value of cell: " << value << '\n';
	
    SysFreeString(bstrValue);
    return value;
}

int main()
{
    CoInitialize(NULL);
    
    IUIAutomation* pAutomation = NULL;
    HRESULT hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, 
                                  __uuidof(IUIAutomation), (void**)&pAutomation);
	std::wcout << L"Value 1\n";
    if (SUCCEEDED(hr))
    {
		std::wcout << L"Value 2\n";
        IUIAutomationElement* pWindow = NULL;
        hr = pAutomation->GetRootElement(&pWindow);

        if (SUCCEEDED(hr) && pWindow)
        {
			std::wcout << L"Value 3\n";

            IUIAutomationElement* pTable = NULL;
            IUIAutomationCondition* pCondition = NULL;

            VARIANT varProp;
            VariantInit(&varProp);
            varProp.vt = VT_BSTR;
            varProp.bstrVal = SysAllocString(L"Table App");
			std::wcout << L"Value 4\n";
            hr = pAutomation->CreatePropertyCondition(UIA_ClassNamePropertyId, varProp, &pCondition);
            std::wcout << L"Value 5\n";
            if (SUCCEEDED(hr))
            {
				std::wcout << L"Value 6\n";
                pWindow->FindFirst(TreeScope_Descendants, pCondition, &pTable);
				std::wcout << L"Value 7\n";
                if (pTable)
                {
					std::wcout << L"Value 8\n";
                    IUIAutomationGridPattern* pGridPattern = NULL;
                    hr = pTable->GetCurrentPattern(UIA_GridPatternId, (IUnknown**)&pGridPattern);
                    std::wcout << L"Value 9\n";
                    if (SUCCEEDED(hr) && pGridPattern)
                    {
						std::wcout << L"Value 10\n";
                        IUIAutomationElement* pCell = NULL;
                        hr = pGridPattern->GetItem(0, 0, &pCell);
                        
                        if (SUCCEEDED(hr) && pCell)
                        {
                            std::wcout << L"Value of cell (0,0): " << GetCellValue(pCell) << std::endl;
                            
                            pCell->Release();
                        }
                        
                        pGridPattern->Release();
                    }
                    
                    pTable->Release();
                }
                
                pCondition->Release();
            }

            VariantClear(&varProp);
            pWindow->Release();
        }
        
        pAutomation->Release();
    }
    
    CoUninitialize();
    return 0;
}
*/

#include <Windows.h>
#include <UIAutomation.h>
#include <iostream>
#include <string>
#include <oleauto.h>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "UIAutomationCore.lib")

std::wstring GetCellValue(IUIAutomationElement* pCell) {
    BSTR bstrValue;
    pCell->get_CurrentName(&bstrValue);
    std::wstring value(bstrValue, SysStringLen(bstrValue));
    
    std::wcout << L"Value of cell: " << value << '\n';
    
    SysFreeString(bstrValue);
    return value;
}

int main()
{
    CoInitialize(NULL);
    
    IUIAutomation* pAutomation = NULL;
    HRESULT hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, 
                                  __uuidof(IUIAutomation), (void**)&pAutomation);
    std::wcout << L"Value 1\n";
    if (SUCCEEDED(hr))
    {
        std::wcout << L"Value 2\n";
        IUIAutomationElement* pRoot = NULL;
        hr = pAutomation->GetRootElement(&pRoot);

        if (SUCCEEDED(hr) && pRoot)
        {
            std::wcout << L"Value 3\n";

            IUIAutomationElement* pWindow = NULL;
            IUIAutomationCondition* pWindowCondition = NULL;

            VARIANT varWindowProp;
            VariantInit(&varWindowProp);
            varWindowProp.vt = VT_BSTR;
            varWindowProp.bstrVal = SysAllocString(L"Table App");

            hr = pAutomation->CreatePropertyCondition(UIA_NamePropertyId, varWindowProp, &pWindowCondition);
            
            if (SUCCEEDED(hr))
            {
                pRoot->FindFirst(TreeScope_Children, pWindowCondition, &pWindow);
                pWindowCondition->Release();
                VariantClear(&varWindowProp);

                if (pWindow)
                {
                    std::wcout << L"Value 4\n";
                    IUIAutomationElement* pTable = NULL;
                    IUIAutomationCondition* pTableCondition = NULL;

                    VARIANT varTableProp;
                    VariantInit(&varTableProp);
                    varTableProp.vt = VT_I4;
                    varTableProp.lVal = UIA_ListControlTypeId;

                    hr = pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, varTableProp, &pTableCondition);
                    std::wcout << L"Value 5\n";
                    if (SUCCEEDED(hr))
                    {
                        std::wcout << L"Value 6\n";
                        pWindow->FindFirst(TreeScope_Descendants, pTableCondition, &pTable);
                        std::wcout << L"Value 7\n";
                        if (pTable)
                        {
                            std::wcout << L"Value 8\n";
                            IUIAutomationGridPattern* pGridPattern = NULL;
                            hr = pTable->GetCurrentPattern(UIA_GridPatternId, (IUnknown**)&pGridPattern);
                            std::wcout << L"Value 9\n";
                            if (SUCCEEDED(hr) && pGridPattern)
                            {
                                std::wcout << L"Value 10\n";
                                IUIAutomationElement* pCell = NULL;
                                hr = pGridPattern->GetItem(0, 0, &pCell);
                                
                                if (SUCCEEDED(hr) && pCell)
                                {
                                    std::wcout << L"Value of cell (0,0): " << GetCellValue(pCell) << std::endl;
                                    
                                    pCell->Release();
                                }
                                
                                pGridPattern->Release();
                            }
                            
                            pTable->Release();
                        }
                        else
                        {
                            std::wcout << L"Table not found\n";
                        }
                        
                        pTableCondition->Release();
                    }

                    VariantClear(&varTableProp);
                    pWindow->Release();
                }
                else
                {
                    std::wcout << L"Window 'Table App' not found\n";
                }
            }

            pRoot->Release();
        }
        
        pAutomation->Release();
    }
    
    CoUninitialize();
    return 0;
}