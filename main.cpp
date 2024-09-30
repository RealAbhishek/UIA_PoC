#include <iostream>
#include <string>
#include <tuple>
#include "UIElement.h"
#include <Windows.h>


void ScrollToBottomByInvokingLineDownButton(UIElement& lineDownButton)
{
    // Get the InvokePattern from the button
    IUIAutomationInvokePattern* pInvokePattern = nullptr;
    HRESULT hr = lineDownButton.GetRawElement()->GetCurrentPatternAs(
        UIA_InvokePatternId, __uuidof(IUIAutomationInvokePattern), (void**)&pInvokePattern);

    if (SUCCEEDED(hr) && pInvokePattern)
    {
        std::wcout << L"Invoking LineDown button to scroll\n";

        // Option 1: Invoke the button a fixed number of times
        const int maxInvocations = 100; // Adjust based on expected content length
        for (int i = 0; i < maxInvocations; ++i)
        {
            hr = pInvokePattern->Invoke();
            if (FAILED(hr))
            {
                std::wcout << L"Failed to invoke LineDown button. HRESULT: " << std::hex << hr << L"\n";
                break;
            }

            // Optionally, introduce a small delay to allow the UI to update
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
    // Get the InvokePattern from the button
    IUIAutomationInvokePattern* pInvokePattern = nullptr;
    HRESULT hr = columnRightButton.GetRawElement()->GetCurrentPatternAs(
        UIA_InvokePatternId, __uuidof(IUIAutomationInvokePattern), (void**)&pInvokePattern);

    if (SUCCEEDED(hr) && pInvokePattern)
    {
        std::wcout << L"Invoking 'Column right' button to scroll\n";

        // Set a reasonable maximum number of invocations
        const int maxInvocations = 100; // Adjust based on expected content width
        for (int i = 0; i < maxInvocations; ++i)
        {
            hr = pInvokePattern->Invoke();
            if (FAILED(hr))
            {
                std::wcout << L"Failed to invoke 'Column right' button. HRESULT: " << std::hex << hr << L"\n";
                break;
            }

            // Optionally, introduce a small delay to allow the UI to update
            Sleep(50);

            // Optionally, check if the button is still enabled
            BOOL isEnabled = FALSE;
            hr = columnRightButton.GetRawElement()->get_CurrentIsEnabled(&isEnabled);
            if (SUCCEEDED(hr) && !isEnabled)
            {
                std::wcout << L"'Column right' button is no longer enabled.\n";
                break;
            }
        }

        pInvokePattern->Release();
        std::wcout << L"Finished invoking 'Column right' button\n";
    }
    else
    {
        std::wcout << L"Failed to get InvokePattern from 'Column right' button. HRESULT: " << std::hex << hr << L"\n";
    }
}

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

            // Spinners Columns and Rows
            auto pSpinner = uiAutomation.FindElementByNameAndType(*pWindow, L"Spinner", UIA_EditControlTypeId);
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

            auto pColumns = uiAutomation.FindElementByNameAndType(*pWindow, L"Columns", UIA_EditControlTypeId);
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
            // Spinners <-------

            // Scrollbar
            // Find vertical scrollbar
            auto verticalScrollBar = uiAutomation.FindScrollBarByName(*pGrid, L"Vertical Scroll Bar");
            if (verticalScrollBar)
            {
                std::wcout << L"Found vertical scrollbar:\n";
                verticalScrollBar->PrintInfo(L"    ");

                // Find the "Line down" button within the vertical scrollbar
                auto pLineDownButton = uiAutomation.FindLineDownButton(*verticalScrollBar);

                if (pLineDownButton)
                {
                    std::wcout << L"Found LineDown button:\n";
                    pLineDownButton->PrintInfo(L"      ");

                    // Invoke the button repeatedly to scroll to the bottom
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

            // Find horizontal scrollbar
            auto horizontalScrollBar = uiAutomation.FindScrollBarByName(*pGrid, L"Horizontal Scroll Bar");
            if (horizontalScrollBar)
            {
                std::wcout << L"Found horizontal scrollbar:\n";
                horizontalScrollBar->PrintInfo(L"    ");


                // Find the "Column right" button within the horizontal scrollbar
                auto pColumnRightButton = uiAutomation.FindColumnRightButton(*horizontalScrollBar);

                if (pColumnRightButton)
                {
                    std::wcout << L"Found 'Column right' button:\n";
                    pColumnRightButton->PrintInfo(L"      ");

                    // Invoke the button repeatedly to scroll to the right
                    ScrollToRightByInvokingColumnRightButton(*pColumnRightButton);
                }
                else
                {
                    std::wcout << L"'Column right' button not found\n";
                }
            }
            else
            {
                std::wcout << L"Horizontal scrollbar not found\n";
            }
            // <----------


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