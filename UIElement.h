#pragma once
#include <Windows.h>
#include <UIAutomation.h>
#include <iostream>
#include <string>
#include <oleauto.h>
#include <memory>
#include <vector>
#include <queue>

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

    std::unique_ptr<UIElement> FindElementByName(const UIElement& parent, const std::wstring& elementName)
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

};

