#include <iostream>
#include <string>
#include <tuple>
#include "UIElement.h"
#include <Windows.h>
#include <chrono>

const std::wstring& length = L"1000";
const int maxInvocations = 1000;

void ScrollToBottomByInvokingLineDownButton(UIElement& lineDownButton)
{
    IUIAutomationInvokePattern* pInvokePattern = nullptr;
    HRESULT hr = lineDownButton.GetRawElement()->GetCurrentPatternAs(
        UIA_InvokePatternId, __uuidof(IUIAutomationInvokePattern), (void**)&pInvokePattern);

    if (SUCCEEDED(hr) && pInvokePattern)
    {
        std::wcout << L"Invoking LineDown button to scroll\n";

        //const int maxInvocations = 1000;
        for (int i = 0; i < maxInvocations; ++i)
        {
            hr = pInvokePattern->Invoke();
            if (FAILED(hr))
            {
                std::wcout << L"Failed to invoke LineDown button. HRESULT: " << std::hex << hr << L"\n";
                break;
            }
            Sleep(50);
        }

        pInvokePattern->Release();
        std::wcout << L"Finished invoking LineDown button\n";
    }
    else
    {
        std::wcout << L"Failed to get InvokePattern from LineDown button. HRESULT: " << std::hex << hr << L"\n";
    }
}

void ScrollToRightByInvokingColumnRightButton(UIElement& columnRightButton)
{
    IUIAutomationInvokePattern* pInvokePattern = nullptr;
    HRESULT hr = columnRightButton.GetRawElement()->GetCurrentPatternAs(
        UIA_InvokePatternId, __uuidof(IUIAutomationInvokePattern), (void**)&pInvokePattern);

    if (SUCCEEDED(hr) && pInvokePattern)
    {
        std::wcout << L"Invoking ColumnRight button to scroll\n";

        //const int maxInvocations = 1000;
        for (int i = 0; i < maxInvocations; ++i)
        {
            hr = pInvokePattern->Invoke();
            if (FAILED(hr))
            {
                std::wcout << L"Failed to invoke ColumnRight button. HRESULT: " << std::hex << hr << L"\n";
                break;
            }

            Sleep(50);

            BOOL isEnabled = FALSE;
            hr = columnRightButton.GetRawElement()->get_CurrentIsEnabled(&isEnabled);
            if (SUCCEEDED(hr) && !isEnabled)
            {
                std::wcout << L"ColumnRight button is no longer enabled.\n";
                break;
            }
        }

        pInvokePattern->Release();
        std::wcout << L"Finished invoking ColumnRight button\n";
    }
    else
    {
        std::wcout << L"Failed to get InvokePattern from ColumnRight button. HRESULT: " << std::hex << hr << L"\n";
    }
}

//int main() 
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
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
            // auto list = uiAutomation.GetElements(*pWindow, children);


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

            auto pSpinner = uiAutomation.FindElementByNameAndType(*pWindow, L"Spinner", UIA_EditControlTypeId);
            if (pSpinner)
            {
                std::wcout << L"Found Spinner element:\n";
                pSpinner->PrintInfo(L"  ");

                IUIAutomationValuePattern* pValuePattern = nullptr;
                HRESULT hr = pSpinner->GetRawElement()->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), (void**)&pValuePattern);

                if (SUCCEEDED(hr) && pValuePattern)
                {
                    BSTR bstrValue = SysAllocString(length.c_str());
                    hr = pValuePattern->SetValue(bstrValue);
                    if (SUCCEEDED(hr))
                    {
                        std::wcout << L"Spinner value set to - " << length << "\n";
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

            auto pColumns = uiAutomation.FindElementByNameAndType(*pWindow, L"Columns", UIA_EditControlTypeId);
            if (pColumns)
            {
                std::wcout << L"Found Spinner element:\n";
                pColumns->PrintInfo(L"  ");

                IUIAutomationValuePattern* pValuePattern = nullptr;
                HRESULT hr = pColumns->GetRawElement()->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), (void**)&pValuePattern);

                if (SUCCEEDED(hr) && pValuePattern)
                {
                    BSTR bstrValue = SysAllocString(length.c_str());
                    hr = pValuePattern->SetValue(bstrValue);
                    if (SUCCEEDED(hr))
                    {
                        std::wcout << L"Spinner value set to - " << length << "\n";
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
            //Sleep(10000);
            ////
            //int targetRow = 995;
            //int targetColumn = 995;
            //auto pCell = uiAutomation.FindCellByRowAndColumn(*pGrid, targetRow, targetColumn);
            //if (pCell)
            //{
            //    std::wcout << L"Found cell:\n";
            //    pCell->PrintInfo(L"    ");

            //    // Set the value in the cell
            //    std::wstring newValue = L"adubey@workfusion.com";
            //    if (uiAutomation.SetCellValue(*pCell, newValue))
            //    {
            //        std::wcout << L"Successfully set value '" << newValue << L"' in cell.\n";
            //    }
            //    else
            //    {
            //        std::wcout << L"Failed to set value in cell.\n";
            //    }
            //    VARIANT varCellvalue;
            //    VariantInit(&varCellvalue);

            //    auto a = pCell->GetName();
            //    auto rhr = pCell->GetRawElement()->GetCurrentPropertyValue(UIA_ValueValuePropertyId, &varCellvalue);
            //    if (SUCCEEDED(rhr) && varCellvalue.vt == VT_BSTR)
            //    {
            //        BSTR result = varCellvalue.bstrVal;
            //        std::wstring value(result, SysStringLen(result));
            //        auto Message = value + L" - Retrieved Without using Scrollbars";
            //        
            //        VariantClear(&varCellvalue);
            //    }
            //    else
            //        VariantClear(&varCellvalue);
            //}
            // 
            // Spinners <-------

            // Scrollbar
            auto verticalScrollBar = uiAutomation.FindScrollBarByName(*pGrid, L"Vertical Scroll Bar");
            //auto verticalScrollBar = uiAutomation.FindScrollBarByNameWithWait(*pGrid, L"Vertical Scroll Bar"); 
            auto start = std::chrono::high_resolution_clock::now();
            if (verticalScrollBar)
            {
                std::wcout << L"Found vertical scrollbar:\n";
                verticalScrollBar->PrintInfo(L"    ");

                auto pLineDownButton = uiAutomation.FindLineDownButton(*verticalScrollBar);

                if (pLineDownButton)
                {
                    std::wcout << L"Found LineDown button:\n";
                    pLineDownButton->PrintInfo(L"      ");

                    ScrollToBottomByInvokingLineDownButton(*pLineDownButton);
                }
                else
                {
                    std::wcout << L"LineDown button not found\n";
                }
            }
            else
            {
                std::wcout << L"Vertical scrollbar not found\n";
            }

            auto horizontalScrollBar = uiAutomation.FindScrollBarByName(*pGrid, L"Horizontal Scroll Bar");
            if (horizontalScrollBar)
            {
                std::wcout << L"Found horizontal scrollbar:\n";
                horizontalScrollBar->PrintInfo(L"    ");

                auto pColumnRightButton = uiAutomation.FindColumnRightButton(*horizontalScrollBar);

                if (pColumnRightButton)
                {
                    std::wcout << L"Found ColumnRight button:\n";
                    pColumnRightButton->PrintInfo(L"      ");

                    ScrollToRightByInvokingColumnRightButton(*pColumnRightButton);
                }
                else
                {
                    std::wcout << L"ColumnRight button not found\n";
                }
            }
            else
            {
                std::wcout << L"Horizontal scrollbar not found\n";
            }
            // <----------

            // Find cell in the grid/table of UIA_EditControlTypeId control type
            int rows = 995;
            int columns = 995;

            auto pCellItem = uiAutomation.FindCellByRowAndColumn(*pGrid, rows, columns);
            if (pCellItem)
            {
                std::wcout << L"Found cell:\n";
                pCellItem->PrintInfo(L"    ");

                std::wstring newValue = L"adubey@workfusion.com";
                if (uiAutomation.SetCellValue(*pCellItem, newValue))
                {
                    std::wcout << L"Successfully set value '" << newValue << L"' in cell.\n";
                }
                else
                {
                    std::wcout << L"Failed to set value in cell.\n";
                }
                VARIANT varCellvalue;
                VariantInit(&varCellvalue);
                
                auto a = pCellItem->GetName();
                auto rhr = pCellItem->GetRawElement()->GetCurrentPropertyValue(UIA_ValueValuePropertyId, &varCellvalue);
                if (SUCCEEDED(rhr) && varCellvalue.vt == VT_BSTR)
                {
                    BSTR result = varCellvalue.bstrVal;
                    std::wstring value(result, SysStringLen(result));
                    
                    VariantClear(&varCellvalue);
                }
                else
                    VariantClear(&varCellvalue);
            }
            // <----------

            auto end = std::chrono::high_resolution_clock::now();

            std::chrono::duration<double, std::milli> elapsed = end - start;
            std::cout << "Function took " << elapsed.count() << " ms to execute.\n";
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