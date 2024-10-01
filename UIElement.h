#pragma once
#include <Windows.h>
#include <UIAutomation.h>
#include <iostream>
#include <string>
#include <oleauto.h>
#include <memory>
#include <vector>
#include <queue>
#include <comutil.h>


// Define OrientationType constants if not defined
#ifndef OrientationType_None
#define OrientationType_None 0
#endif

#ifndef OrientationType_Horizontal
#define OrientationType_Horizontal 1
#endif

#ifndef OrientationType_Vertical
#define OrientationType_Vertical 2
#endif

#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "UIAutomationCore.lib")

class UIElement 
{
private:
    IUIAutomationElement* m_pElement;

public:
    UIElement(IUIAutomationElement* pElement) : m_pElement(pElement) 
    {
        std::wcout << L"UIElement Constructor\n";
    }

    ~UIElement() 
    {
        if (m_pElement) 
        {
            std::wcout << L"UIElement Destructor. Release memory held by m_pElement.\n";
            m_pElement->Release();
        }
    }

    std::wstring GetName() const 
    {
        std::wcout << L"Get Name\n";
        BSTR bstrName;
        m_pElement->get_CurrentName(&bstrName);
        std::wstring name(bstrName, SysStringLen(bstrName));
        SysFreeString(bstrName);

        std::wcout << L"Name: " << name << L'\n';
        return name;
    }

    std::wstring GetClassName() const 
    {
        std::wcout << L"Get Class Name.\n";
        BSTR bstrClassName;
        m_pElement->get_CurrentClassName(&bstrClassName);
        std::wstring className(bstrClassName, SysStringLen(bstrClassName));
        SysFreeString(bstrClassName);

        std::wcout << L"Class name: " << className << L'\n';
        return className;
    }

    int GetControlType() const 
    {
        std::wcout << "Get Control Type" << L'\n';
        int controlType;
        m_pElement->get_CurrentControlType(&controlType);

        std::wcout << "Control Type: " << controlType << L'\n';
        return controlType;
    }

    void PrintInfo(const std::wstring& prefix) const 
    {
        std::wcout << prefix << L"Name: " << GetName() << std::endl;
        std::wcout << prefix << L"ClassName: " << GetClassName() << std::endl;
        std::wcout << prefix << L"ControlType: " << GetControlType() << std::endl;
    }

    IUIAutomationElement* GetRawElement() const 
    {
        std::wcout << L"Get  Raw Element\n";
        return m_pElement;
    }
};

class UIAutomationWrapper 
{
private:
    IUIAutomation* m_pAutomation;

public:
    UIAutomationWrapper() : m_pAutomation(nullptr) 
    {
        std::wcout << L"UIAutomationWrapper contructor. Initializing dreadful COM\n";
        CoInitialize(NULL);
        HRESULT hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER,
            __uuidof(IUIAutomation), (void**)&m_pAutomation);
        if (FAILED(hr)) 
        {
            throw std::runtime_error("Failed to create UI Automation instance");
        }
    }

    ~UIAutomationWrapper() 
    {
        std::wcout << L"UIAutomationWrapper destructor. Releasing COM\n";
        if (m_pAutomation) 
        {
            m_pAutomation->Release();
        }
        CoUninitialize();
    }

    std::unique_ptr<UIElement> GetRootElement() 
    {
        std::wcout << L"GetRootElement\n";
        IUIAutomationElement* pRoot = nullptr;
        HRESULT hr = m_pAutomation->GetRootElement(&pRoot);
        if (SUCCEEDED(hr) && pRoot) 
        {
            return std::make_unique<UIElement>(pRoot);
        }
        return nullptr;
    }

    std::unique_ptr<UIElement> FindWindowByName(const std::wstring& windowName) 
    {
        std::wcout << L"GetRootElement\n";
        auto pRoot = GetRootElement();
        if (!pRoot) return nullptr;

        IUIAutomationCondition* pWindowCondition = nullptr;
        VARIANT varWindowProp;
        VariantInit(&varWindowProp);
        varWindowProp.vt = VT_BSTR;
        varWindowProp.bstrVal = SysAllocString(windowName.c_str());

        HRESULT hr = m_pAutomation->CreatePropertyCondition(UIA_NamePropertyId, varWindowProp, &pWindowCondition);
        VariantClear(&varWindowProp);

        if (SUCCEEDED(hr)) 
        {
            IUIAutomationElement* pWindow = nullptr;
            pRoot->GetRawElement()->FindFirst(TreeScope_Children, pWindowCondition, &pWindow);
            pWindowCondition->Release();

            if (pWindow) 
            {
                return std::make_unique<UIElement>(pWindow);
            }
        }

        return nullptr;
    }

    std::unique_ptr<UIElement> FindGridElement(const UIElement& parent) 
    {
        std::wcout << L"FindGridElement\n";
        IUIAutomationCondition* pCondition = nullptr;
        IUIAutomationElement* pGrid = nullptr;

        VARIANT varProp;
        VariantInit(&varProp);
        varProp.vt = VT_BSTR;
        varProp.bstrVal = SysAllocString(L"DataGridView");

        HRESULT hr = m_pAutomation->CreatePropertyCondition(UIA_ClassNamePropertyId, varProp, &pCondition);
        std::wcout << L"FindGridElement A\n";
        if (SUCCEEDED(hr))
        {
            std::wcout << L"FindGridElement B\n";
            parent.GetRawElement()->FindFirst(TreeScope_Descendants, pCondition, &pGrid);
            pCondition->Release();
        }

        VariantClear(&varProp);

        if (!pGrid) 
        {
            std::wcout << L"FindGridElement C\n";
            varProp.vt = VT_I4;
            varProp.lVal = UIA_DataGridControlTypeId;
            hr = m_pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, varProp, &pCondition);
            if (SUCCEEDED(hr)) 
            {
                std::wcout << L"FindGridElement D\n";
                parent.GetRawElement()->FindFirst(TreeScope_Descendants, pCondition, &pGrid);
                pCondition->Release();
            }
        }

        if (!pGrid)
        {
            BSTR bstrName;

            std::wcout << L"Create condition with UIA_TableControlTypeId to get table." << std::endl;
            varProp.vt = VT_I4;
            varProp.lVal = UIA_TableControlTypeId;
            hr = m_pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, varProp, &pCondition);
            if (SUCCEEDED(hr))
            {
                hr = parent.GetRawElement()->FindFirst(TreeScope_Descendants, pCondition, &pGrid);
                if (SUCCEEDED(hr))
                {
                    std::wcout << L"Found Table: " << pGrid->get_CurrentName(&bstrName) << std::endl;

                    std::wstring elementName(bstrName, SysStringLen(bstrName));
                    std::wcout << L"Element name (table name): " << elementName << L'\n';
                    SysFreeString(bstrName);
                    pCondition->Release();
                }
                else
                {
                    pCondition->Release();
                }
            }
        }

        VariantClear(&varProp);

        if (pGrid) 
        {
            std::wcout << L"FindGridElement D\n";
            return std::make_unique<UIElement>(pGrid);
        }

        return nullptr;
    }

    std::unique_ptr<UIElement> FindSpinnerElement(const UIElement& parent)
    {
        std::wcout << L"Finding Spinner element\n";
        IUIAutomationCondition* pNameCondition = nullptr;
        IUIAutomationCondition* pControlTypeCondition = nullptr;
        IUIAutomationCondition* pAndCondition = nullptr;
        IUIAutomationElement* pSpinnerElement = nullptr;
        VARIANT varProp;
        VariantInit(&varProp);

        varProp.vt = VT_BSTR;
        varProp.bstrVal = SysAllocString(L"Spinner");
        HRESULT hr = m_pAutomation->CreatePropertyCondition(UIA_NamePropertyId, varProp, &pNameCondition);
        VariantClear(&varProp);

        if (SUCCEEDED(hr))
        {
            varProp.vt = VT_I4;
            varProp.lVal = UIA_EditControlTypeId;
            hr = m_pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, varProp, &pControlTypeCondition);
            VariantClear(&varProp);
        }

        if (SUCCEEDED(hr))
        {
            hr = m_pAutomation->CreateAndCondition(pNameCondition, pControlTypeCondition, &pAndCondition);
        }

        if (SUCCEEDED(hr))
        {
            hr = parent.GetRawElement()->FindFirst(TreeScope_Descendants, pAndCondition, &pSpinnerElement);
        }

        if (pNameCondition) pNameCondition->Release();
        if (pControlTypeCondition) pControlTypeCondition->Release();
        if (pAndCondition) pAndCondition->Release();

        if (SUCCEEDED(hr) && pSpinnerElement)
        {
            return std::make_unique<UIElement>(pSpinnerElement);
        }
        else
        {
            if (pSpinnerElement) pSpinnerElement->Release();
            return nullptr;
        }
    }

    std::unique_ptr<UIElement> FindElementByNameAndType(const UIElement& parent, const std::wstring& elementName, const long& type)
    {

        std::wcout << L"Finding Spinner element\n";
        IUIAutomationCondition* pNameCondition = nullptr;
        IUIAutomationCondition* pControlTypeCondition = nullptr;
        IUIAutomationCondition* pAndCondition = nullptr;
        IUIAutomationElement* pSpinnerElement = nullptr;
        VARIANT varProp;
        VariantInit(&varProp);

        varProp.vt = VT_BSTR;
        varProp.bstrVal = SysAllocString(elementName.c_str());
        HRESULT hr = m_pAutomation->CreatePropertyCondition(UIA_NamePropertyId, varProp, &pNameCondition);
        VariantClear(&varProp);

        if (SUCCEEDED(hr))
        {
            varProp.vt = VT_I4;
            varProp.lVal = type;// UIA_EditControlTypeId;
            hr = m_pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, varProp, &pControlTypeCondition);
            VariantClear(&varProp);
        }

        if (SUCCEEDED(hr))
        {
            hr = m_pAutomation->CreateAndCondition(pNameCondition, pControlTypeCondition, &pAndCondition);
        }

        if (SUCCEEDED(hr))
        {
            hr = parent.GetRawElement()->FindFirst(TreeScope_Descendants, pAndCondition, &pSpinnerElement);
        }

        if (pNameCondition) pNameCondition->Release();
        if (pControlTypeCondition) pControlTypeCondition->Release();
        if (pAndCondition) pAndCondition->Release();

        if (SUCCEEDED(hr) && pSpinnerElement)
        {
            return std::make_unique<UIElement>(pSpinnerElement);
        }
        else
        {
            if (pSpinnerElement) pSpinnerElement->Release();
            return nullptr;
        }
    }

    std::unique_ptr<UIElement> FindScrollBarByName(const UIElement& parent, const std::wstring& scrollbarName)
    {
        std::wcout << L"Finding scrollbar with name: " << scrollbarName << L'\n';
        IUIAutomationElement* pScrollBarElement = nullptr;
        IUIAutomationCondition* pScrollBarCondition = nullptr;
        IUIAutomationCondition* pNameCondition = nullptr;
        IUIAutomationCondition* pAndCondition = nullptr;
        HRESULT hr = S_OK;

        hr = m_pAutomation->CreatePropertyCondition(
            UIA_ControlTypePropertyId,
            _variant_t(UIA_ScrollBarControlTypeId),
            &pScrollBarCondition);

        if (SUCCEEDED(hr))
        {
            hr = m_pAutomation->CreatePropertyCondition(
                UIA_NamePropertyId,
                _variant_t(scrollbarName.c_str()),
                &pNameCondition);
        }

        if (SUCCEEDED(hr))
        {
            hr = m_pAutomation->CreateAndCondition(pScrollBarCondition, pNameCondition, &pAndCondition);
        }

        if (SUCCEEDED(hr))
        {
            hr = parent.GetRawElement()->FindFirst(TreeScope_Subtree, pAndCondition, &pScrollBarElement);
        }

        if (pScrollBarCondition) pScrollBarCondition->Release();
        if (pNameCondition) pNameCondition->Release();
        if (pAndCondition) pAndCondition->Release();

        if (SUCCEEDED(hr) && pScrollBarElement)
        {
            return std::make_unique<UIElement>(pScrollBarElement);
        }
        else
        {
            if (pScrollBarElement) pScrollBarElement->Release();
            return nullptr;
        }
    }

    std::unique_ptr<UIElement> FindScrollBarByNameWithWait(const UIElement& parent, const std::wstring& scrollbarName, int maxRetries = 10, int delayMs = 100)
    {
        std::wcout << L"Attempting to find scrollbar with name: " << scrollbarName << L'\n';
        std::unique_ptr<UIElement> scrollbar = nullptr;

        for (int i = 0; i < maxRetries; ++i)
        {
            scrollbar = FindScrollBarByName(parent, scrollbarName);
            if (scrollbar)
            {
                std::wcout << L"Scrollbar found after " << (i + 1) << " attempt(s).\n";
                return scrollbar;
            }

            std::wcout << L"Scrollbar not found. Waiting for " << delayMs << L" ms.\n";
            Sleep(delayMs);
        }

        std::wcout << L"Scrollbar not found after " << maxRetries << " attempts.\n";
        return nullptr;
    }

    std::unique_ptr<UIElement> FindLineDownButton(const UIElement& parent)
    {
        std::wcout << L"Finding LineDown button\n";
        IUIAutomationElement* pButtonElement = nullptr;
        IUIAutomationCondition* pButtonCondition = nullptr;
        IUIAutomationCondition* pNameCondition = nullptr;
        IUIAutomationCondition* pAndCondition = nullptr;
        HRESULT hr = S_OK;

        hr = m_pAutomation->CreatePropertyCondition(
            UIA_ControlTypePropertyId,
            _variant_t(UIA_ButtonControlTypeId),
            &pButtonCondition);

        if (SUCCEEDED(hr))
        {
            hr = m_pAutomation->CreatePropertyCondition(
                UIA_NamePropertyId,
                _variant_t(L"Line down"),
                &pNameCondition);
        }

        if (SUCCEEDED(hr))
        {
            hr = m_pAutomation->CreateAndCondition(pButtonCondition, pNameCondition, &pAndCondition);
        }

        if (SUCCEEDED(hr))
        {
            hr = parent.GetRawElement()->FindFirst(TreeScope_Subtree, pAndCondition, &pButtonElement);
        }

        if (pButtonCondition) pButtonCondition->Release();
        if (pNameCondition) pNameCondition->Release();
        if (pAndCondition) pAndCondition->Release();

        if (SUCCEEDED(hr) && pButtonElement)
        {
            return std::make_unique<UIElement>(pButtonElement);
        }
        else
        {
            if (pButtonElement) pButtonElement->Release();
            return nullptr;
        }
    }

    std::unique_ptr<UIElement> FindColumnRightButton(const UIElement& parent)
    {
        std::wcout << L"Finding 'Column right' button\n";
        IUIAutomationElement* pButtonElement = nullptr;
        IUIAutomationCondition* pButtonCondition = nullptr;
        IUIAutomationCondition* pNameCondition = nullptr;
        IUIAutomationCondition* pAndCondition = nullptr;
        HRESULT hr = S_OK;

        hr = m_pAutomation->CreatePropertyCondition(
            UIA_ControlTypePropertyId,
            _variant_t(UIA_ButtonControlTypeId),
            &pButtonCondition);

        if (SUCCEEDED(hr))
        {
            hr = m_pAutomation->CreatePropertyCondition(
                UIA_NamePropertyId,
                _variant_t(L"Column right"),
                &pNameCondition);
        }

        if (SUCCEEDED(hr))
        {
            hr = m_pAutomation->CreateAndCondition(pButtonCondition, pNameCondition, &pAndCondition);
        }

        if (SUCCEEDED(hr))
        {
            hr = parent.GetRawElement()->FindFirst(TreeScope_Subtree, pAndCondition, &pButtonElement);
        }

        if (pButtonCondition) pButtonCondition->Release();
        if (pNameCondition) pNameCondition->Release();
        if (pAndCondition) pAndCondition->Release();

        if (SUCCEEDED(hr) && pButtonElement)
        {
            return std::make_unique<UIElement>(pButtonElement);
        }
        else
        {
            if (pButtonElement) pButtonElement->Release();
            return nullptr;
        }
    }

    void ScrollToBottomByInvokingLineDownButton(UIElement& lineDownButton)
    {
        IUIAutomationInvokePattern* pInvokePattern = nullptr;
        HRESULT hr = lineDownButton.GetRawElement()->GetCurrentPatternAs(
            UIA_InvokePatternId, __uuidof(IUIAutomationInvokePattern), (void**)&pInvokePattern);

        if (SUCCEEDED(hr) && pInvokePattern)
        {
            std::wcout << L"Invoking LineDown button to scroll\n";

            const int maxInvocations = 100;
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

    std::unique_ptr<UIElement> FindCellByRowAndColumn(const UIElement& gridElement, int row, int column)
    {
        std::wcout << L"Finding cell at Row: " << row << L", Column: " << column << L'\n';

        IUIAutomationCondition* pCondition = nullptr;
        HRESULT hr = m_pAutomation->CreatePropertyCondition(
            UIA_ControlTypePropertyId,
            _variant_t(UIA_EditControlTypeId),
            &pCondition);

        if (SUCCEEDED(hr))
        {
            IUIAutomationElementArray* pFoundElements = nullptr;
            hr = gridElement.GetRawElement()->FindAll(TreeScope_Subtree, pCondition, &pFoundElements);
            pCondition->Release();

            if (SUCCEEDED(hr) && pFoundElements)
            {
                int length = 0;
                pFoundElements->get_Length(&length);
                for (int i = 0; i < length; ++i)
                {
                    IUIAutomationElement* pElement = nullptr;
                    hr = pFoundElements->GetElement(i, &pElement);
                    if (SUCCEEDED(hr) && pElement)
                    {
                        // Check if the element supports GridItemPattern
                        BOOL supportsPattern = FALSE;
                        VARIANT var;
                        VariantInit(&var);
                        //hr = pElement->get_CurrentIsPatternAvailable(UIA_GridItemPatternId, &var);

                        hr = pElement->GetCurrentPropertyValue(UIA_IsGridItemPatternAvailablePropertyId, &var);

                        if (SUCCEEDED(hr) && var.vt == VT_BOOL && var.boolVal == VARIANT_TRUE)
                        {
                            VariantClear(&var);
                            IUIAutomationGridItemPattern* pGridItemPattern = nullptr;
                            hr = pElement->GetCurrentPatternAs(
                                UIA_GridItemPatternId, __uuidof(IUIAutomationGridItemPattern), (void**)&pGridItemPattern);

                            if (SUCCEEDED(hr) && pGridItemPattern)
                            {
                                int cellRow = 0, cellColumn = 0;
                                hr = pGridItemPattern->get_CurrentRow(&cellRow);
                                hr = pGridItemPattern->get_CurrentColumn(&cellColumn);
                                pGridItemPattern->Release();

                                if (cellRow == row && cellColumn == column)
                                {
                                    pFoundElements->Release();
                                    std::wcout << L"Found cell at (" << cellRow << L"," << cellColumn << L")\n";
                                    return std::make_unique<UIElement>(pElement);
                                }
                            }
                        }
                        else
                        {
                            VariantClear(&var);
                        }
                        pElement->Release();
                    }
                }
                pFoundElements->Release();
            }
        }
        else
        {
            std::wcout << L"Failed to create condition for finding cells. HRESULT: " << std::hex << hr << L'\n';
        }

        std::wcout << L"Cell not found at Row: " << row << L", Column: " << column << L'\n';
        return nullptr;
    }

    bool SetCellValue(UIElement& cellElement, const std::wstring& value)
    {
        // Try to get ValuePattern
        IUIAutomationValuePattern* pValuePattern = nullptr;
        HRESULT hr = cellElement.GetRawElement()->GetCurrentPatternAs(
            UIA_ValuePatternId, __uuidof(IUIAutomationValuePattern), (void**)&pValuePattern);

        if (SUCCEEDED(hr) && pValuePattern)
        {
            BSTR bstrValue = SysAllocString(value.c_str());
            hr = pValuePattern->SetValue(bstrValue);
            pValuePattern->Release();
            SysFreeString(bstrValue);
            if (SUCCEEDED(hr))
            {
                std::wcout << L"Value set successfully using ValuePattern\n";
                return true;
            }
            else
            {
                std::wcout << L"Failed to set value using ValuePattern. HRESULT: " << std::hex << hr << L"\n";
                return false;
            }
        }
        else
        {
            std::wcout << L"ValuePattern not available or failed to retrieve. HRESULT: " << std::hex << hr << L"\n";
            return false;
        }
    }
};

