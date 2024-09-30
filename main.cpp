#include <iostream>
#include <string>
#include "UIElement.h"

int main() 
{
    try 
    {
        UIAutomationWrapper uiAutomation;
        IUIAutomationElementArray* children = nullptr;

        auto pWindow = uiAutomation.FindWindowByName(L"Grid Control Demo");
        if (pWindow) 
        {
            std::wcout << L"Found window:\n";
            pWindow->PrintInfo(L"  ");
            // First lets iterate all the elements and creates a list
            // auto list = uiAutomation.GetElements(*pWindow, children); // Issue


            auto pGrid = uiAutomation.FindGridElement(*pWindow);
            if (pGrid) 
            {
                std::wcout << L"Found grid element:\n";
                pGrid->PrintInfo(L"  ");
            } 
            
            else 
            {
                std::wcout << L"Grid element not found\n";
            }

            auto pSpinner = uiAutomation.FindElementByName(*pWindow, L"Spinner");
            if (pSpinner)
            {
                std::wcout << L"Found Spinner element:\n";
                pSpinner->PrintInfo(L"  ");

                // Get the ValuePattern of the Spinner element
                IUIAutomationValuePattern* pValuePattern = nullptr;
                HRESULT hr = pSpinner->GetRawElement()->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), (void**)&pValuePattern);

                if (SUCCEEDED(hr) && pValuePattern)
                {
                    BSTR bstrValue = SysAllocString(L"100");
                    hr = pValuePattern->SetValue(bstrValue);
                    if (SUCCEEDED(hr))
                    {
                        std::wcout << L"Spinner value set to 100\n";
                    }
                    else
                    {
                        std::wcout << L"Failed to set Spinner value\n";
                    }
                    pValuePattern->Release();
                }
                else
                {
                    std::wcout << L"Failed to get ValuePattern from Spinner element\n";
                }
            }

            auto pColumns = uiAutomation.FindElementByName(*pWindow, L"Columns");
            if (pColumns)
            {
                std::wcout << L"Found Spinner element:\n";
                pColumns->PrintInfo(L"  ");

                // Get the ValuePattern of the Spinner element
                IUIAutomationValuePattern* pValuePattern = nullptr;
                HRESULT hr = pColumns->GetRawElement()->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), (void**)&pValuePattern);

                if (SUCCEEDED(hr) && pValuePattern)
                {
                    BSTR bstrValue = SysAllocString(L"100");
                    hr = pValuePattern->SetValue(bstrValue);
                    if (SUCCEEDED(hr))
                    {
                        std::wcout << L"Spinner value set to 100\n";
                    }
                    else
                    {
                        std::wcout << L"Failed to set Spinner value\n";
                    }
                    pValuePattern->Release();
                }
                else
                {
                    std::wcout << L"Failed to get ValuePattern from Spinner element\n";
                }
            }

            else
            {
                std::wcout << L"Spinner element not found\n";
            }

        }
        
        else 
        {
            std::wcout << L"Window 'Grid Control Demo' not found\n";
        }
    } 
    
    catch (const std::exception& e) 
    {
        std::cerr << "Error: " << e.what() << std::endl;
        return 1;
    }

    return 0;
}