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
            auto list = uiAutomation.GetElements(*pWindow, children); // Issue
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
        } 
        
        else 
        {
            std::wcout << L"Window 'Grid Control Demo' not found\n";
        }
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        return 1;
    }

    return 0;
}