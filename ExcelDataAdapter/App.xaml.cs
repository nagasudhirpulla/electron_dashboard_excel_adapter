using AdapterUtils;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelDataAdapter
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            AdapterParams prms = AdapterParams.ParseArgs(e.Args);
            if (prms.RequestType == RequestTypes.PickMeas)
            {
                MeasPickerWindow win = new MeasPickerWindow();
                win.Show();
            }
            else
            {
                // we assume that the request is getting data
                DataFetcher.FetchAndFlushData(prms);
                Current.Shutdown();
            }
        }
    }
}
