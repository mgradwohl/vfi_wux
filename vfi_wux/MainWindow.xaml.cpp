#include "pch.h"
#include "MainWindow.xaml.h"
#if __has_include("MainWindow.g.cpp")
#include "MainWindow.g.cpp"
#endif
#include "mylist.h"
#include <iostream>
#include <locale>

using namespace winrt;
using namespace Microsoft::UI::Xaml;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace winrt::vfi_wux::implementation
{
    MainWindow::MainWindow()
    {
        InitializeComponent();
    }

    int32_t MainWindow::MyProperty()
    {
        throw hresult_not_implemented();
    }

    void MainWindow::MyProperty(int32_t /* value */)
    {
        throw hresult_not_implemented();
    }
}


void winrt::vfi_wux::implementation::MainWindow::AppBarButton_Click(winrt::Windows::Foundation::IInspectable const& sender, winrt::Microsoft::UI::Xaml::RoutedEventArgs const& e)
{
    MyList thelist;
    auto p = std::make_shared<CWiseFile>(L"c:\\windows\\system32\\wwahost.exe");
    thelist.AddHead(p);
    p->ReadVersionInfo();
    p->GetAttribs();
    p->GetCodePage();
    p->SetVersionStrings();
    p->GetProductVersion();
    p->GetFileVersion();
    p->GetFlags();
    p->GetOS();
    p->GetSize();
    p->GetType();


}