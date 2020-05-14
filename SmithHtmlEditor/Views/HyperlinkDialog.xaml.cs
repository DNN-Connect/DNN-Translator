using System.Windows;

namespace Smith.WPF.HtmlEditor
{
    /// <summary>
    /// HyperlinkDialog.xaml 的交互逻辑
    /// </summary>
    public partial class HyperlinkDialog : Window
    {
        HyperlinkObject bindingContext;

        public HyperlinkDialog()
        {
            InitializeComponent();

            Model = new HyperlinkObject
            {
                URL = "http://"
            };
            OkayButton.Click += new RoutedEventHandler(OkayButton_Click);
            CancelButton.Click += new RoutedEventHandler(CancelButton_Click);
        }

        public HyperlinkObject Model
        {
            get { return bindingContext; }
            set
            {
                bindingContext = value;
                this.DataContext = bindingContext;
            }
        }

        void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
            this.Close();
        }

        void OkayButton_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
            if (System.Windows.Interop.ComponentDispatcher.IsThreadModal) this.DialogResult = true;
            this.Close();
        }
    }
}
