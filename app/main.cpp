#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Microsoft.UI.Xaml.Controls.h>
#include <winrt/Microsoft.UI.Xaml.XamlTypeInfo.h>
#include <winrt/Microsoft.UI.Xaml.Markup.h>
#include <winrt/Microsoft.UI.Xaml.h>
#include <winrt/Microsoft.UI.Xaml.Controls.Primitives.h>
#include <winrt/Microsoft.UI.Xaml.Data.h>
#include <winrt/Microsoft.UI.Xaml.Interop.h>
#include <winrt/Microsoft.UI.Xaml.Media.h>

using namespace winrt;
using namespace Microsoft::UI::Xaml;
using namespace Microsoft::UI::Xaml::Controls;
using namespace Microsoft::UI::Xaml::XamlTypeInfo;
using namespace Microsoft::UI::Xaml::Markup;

struct MainWindow : implements<MainWindow, IXamlMetadataProvider>
{
    Window window{ nullptr };
    Grid rootGrid{ nullptr };
    DataGrid dataGrid{ nullptr };

    MainWindow()
    {
        window = Window();
        rootGrid = Grid();
        dataGrid = DataGrid();

        // Configure the DataGrid
        dataGrid.AutoGenerateColumns(true);
        dataGrid.IsReadOnly(false);

        // Add some sample data
        auto items = winrt::single_threaded_observable_vector<IInspectable>();
        for (int i = 0; i < 5; i++)
        {
            auto item = PropertySet();
            item.Insert(L"Column1", box_value(L"Row " + to_hstring(i + 1)));
            item.Insert(L"Column2", box_value(L"Editable"));
            item.Insert(L"Column3", box_value(L"Editable"));
            items.Append(item);
        }
        dataGrid.ItemsSource(items);

        // Add the DataGrid to the root Grid
        rootGrid.Children().Append(dataGrid);

        // Set the window content
        window.Content(rootGrid);
        window.Activate();
    }

    IXamlType GetXamlType(TypeName const&) const noexcept
    {
        return nullptr;
    }

    IXamlType GetXamlType(hstring const&) const noexcept
    {
        return nullptr;
    }

    com_array<XmlnsDefinition> GetXmlnsDefinitions() const noexcept
    {
        return {};
    }
};

int __stdcall wWinMain(HINSTANCE, HINSTANCE, PWSTR, int)
{
    init_apartment(apartment_type::single_threaded);
    Application::Start([](auto&&) {
        make<MainWindow>();
    });
}