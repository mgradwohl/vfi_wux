#pragma once

#include "MainWindow.g.h"
#include "MyList.h"

namespace winrt::vfi_wux::implementation
{
    struct MainWindow : MainWindowT<MainWindow>
    {
    public:
        MainWindow();

        int32_t MyProperty();
        void MyProperty(int32_t value);

        void AppBarButton_Click(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::RoutedEventArgs const& e);
    };
}

namespace winrt::vfi_wux::factory_implementation
{
    struct MainWindow : MainWindowT<MainWindow, implementation::MainWindow>
    {
    };
}
