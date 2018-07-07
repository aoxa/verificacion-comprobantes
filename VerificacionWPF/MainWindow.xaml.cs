using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.ComponentModel;
using VerificacionComprobantes;
using Fizzler.Systems.HtmlAgilityPack;

namespace VerificacionWPF
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();

            VerificacionComprobantes.Service.ConfigurationServer configurationServer = VerificacionComprobantes.Service.ConfigurationServer.Instance;

            string dbConfig = System.IO.Directory.GetCurrentDirectory() + System.IO.Path.DirectorySeparatorChar + "appContent.db";
            // TODO agregar en la app la direcion de la DB
            configurationServer.dbConfig = dbaddress.Text = dbConfig;

            solapas.Text = XLSDataSource.ColumnMappings.SHEETS;
            fecha.Text = XLSDataSource.ColumnMappings.FECHA;
            noperacion.Text = XLSDataSource.ColumnMappings.OPERATION;
            voucher.Text = XLSDataSource.ColumnMappings.VOUCHER;
            nombre.Text = XLSDataSource.ColumnMappings.NAME;
            total.Text = XLSDataSource.ColumnMappings.TOTAL;

            updateConfigurationScreen(configurationServer.getConfiguration());
        }

        private void updateConfigurationScreen(VerificacionComprobantes.Service.ConfigurationServer.Configuration config)
        {
            if (config == null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(config.Email.Server))
                smtp.Text = config.Email.Server;
            if (!string.IsNullOrEmpty(config.Email.Port))
                port.Text = config.Email.Port;
            if (!string.IsNullOrEmpty(config.Email.User))
                user.Text = config.Email.User;
            if (!string.IsNullOrEmpty(config.Email.Password))
                pass.Password = config.Email.Password;
            if (!string.IsNullOrEmpty(config.Email.Header))
                header.Text = config.Email.Header;
            if (!string.IsNullOrEmpty(config.Email.Sender))
                remitente.Text = config.Email.Sender;
            if (!string.IsNullOrEmpty(config.Email.CC))
                ccEmails.Text = config.Email.CC;
            if (!string.IsNullOrEmpty(config.Email.Template))
                templateChooser.Text = config.Email.Template;

            if (!string.IsNullOrEmpty(config.Sheet.Sheets))
                XLSDataSource.ColumnMappings.SHEETS = solapas.Text = config.Sheet.Sheets;
            if (!string.IsNullOrEmpty(config.Sheet.Date))
                XLSDataSource.ColumnMappings.FECHA = fecha.Text = config.Sheet.Date;
            if (!string.IsNullOrEmpty(config.Sheet.Operation))
                XLSDataSource.ColumnMappings.OPERATION = noperacion.Text = config.Sheet.Operation;
            if (!string.IsNullOrEmpty(config.Sheet.Voucher))
                XLSDataSource.ColumnMappings.VOUCHER = voucher.Text = config.Sheet.Voucher;
            if (!string.IsNullOrEmpty(config.Sheet.Name))
                XLSDataSource.ColumnMappings.NAME = nombre.Text = config.Sheet.Name;
            if (!string.IsNullOrEmpty(config.Sheet.Total))
                XLSDataSource.ColumnMappings.TOTAL = total.Text = config.Sheet.Total;

        }

        private void SearchAction(object sender, RoutedEventArgs e)
        {
            comprobantes.Document.Blocks.Clear();
            foundPerson.UnselectAll();
            foundPerson.Items.Clear();

            VerificacionComprobantes.Service.StorageService<VerificacionComprobantes.Model.Person> service = new VerificacionComprobantes.Service.StorageService<VerificacionComprobantes.Model.Person>();
            
            foreach (VerificacionComprobantes.Model.Person person in service.Query(LiteDB.Query.StartsWith("Name", search.Text), "personas"))
            {
                ListBoxItem item = new ListBoxItem();
                item.Content = person.Name;
                
                foundPerson.Items.Add(item);
            }
        }

        private void SelectPerson(object sender, SelectionChangedEventArgs e)
        {
            comprobantes.Document.Blocks.Clear();

            BackgroundWorker worker = new BackgroundWorker();

            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_DoWork;
            worker.ProgressChanged += worker_ProgressChanged;

            worker.RunWorkerAsync();
            
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            string userName = "";
            this.Dispatcher.Invoke(() => {
                sendEmailButton.IsEnabled = false;

                ListBoxItem selected = null;
           
                selected = (ListBoxItem)foundPerson.SelectedItem;

                userName = selected == null ? null : selected.Content.ToString();
                userEmail.Text = "";
            });

            if (null == userName) return;
            
            string tempBox = "";

            VerificacionComprobantes.Service.StorageService<VerificacionComprobantes.Model.Person> service = new VerificacionComprobantes.Service.StorageService<VerificacionComprobantes.Model.Person>();

            VerificacionComprobantes.Model.Person person = service.Query(LiteDB.Query.EQ("Name", userName), "personas").First();
            var index = 0;
                
            foreach (VerificacionComprobantes.Model.Operation operation in person.Operations)
            {
                (sender as BackgroundWorker).ReportProgress((++index*100)/person.Operations.Count);
                
                tempBox += (operation.Entity + " " + operation.FacturationDate + " ");
                tempBox += (operation.OperationNumber + " $" + operation.Total);

                if (!string.IsNullOrEmpty(operation.Voucher))
                {
                    tempBox +=  (" " + operation.Voucher);
                } else if(!string.IsNullOrEmpty(person.Email))
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        userEmail.Text = person.Email;
                        sendEmailButton.IsEnabled = true;
                    });
                }

                tempBox += (Environment.NewLine);
            }

            this.Dispatcher.Invoke(() =>
            {
                comprobantes.AppendText( tempBox);
            });

            (sender as BackgroundWorker).ReportProgress(0);
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.Dispatcher.Invoke(() =>
            {
                sbar.Value = e.ProgressPercentage;
            });
        }

        private void loadTemplate_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();

            if (openFileDialog.ShowDialog() == true)
            {
                templateChooser.Text = openFileDialog.FileName;
            }   
        }

        private void loadEmails_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();

            if (openFileDialog.ShowDialog() == true)
            {
                XLSDataSource.Instance.LoadEmails(openFileDialog.FileName);
            }
        }

        void worker_InitializeDatasource(object sender, DoWorkEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    string filename = openFileDialog.FileName;
                    XLSDataSource dataSource = XLSDataSource.Instance;

                    dataSource.Init(filename, sender as BackgroundWorker);
                }
                catch (VerificacionComprobantes.Exceptions.SheetNotPresentException ex)
                {

                }

                (sender as BackgroundWorker).ReportProgress(0);
            }
        }

        private void updateUsers_Click(object sender, RoutedEventArgs e)
        {       
            BackgroundWorker worker = new BackgroundWorker();

            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_InitializeDatasource;
            worker.ProgressChanged += worker_ProgressChanged;

            worker.RunWorkerAsync();    
        }

        private void sendEmail(object sender, RoutedEventArgs e)
        {
            VerificacionComprobantes.Service.ConfigurationServer configServer = VerificacionComprobantes.Service.ConfigurationServer.Instance;
            VerificacionComprobantes.Service.ConfigurationServer.Configuration config = configServer.getConfiguration();

            if (string.IsNullOrEmpty(config.Email.Server))
            {
                //logger.Buffer.Text += "No esta configurado el servidor de emails\n";
                return;
            }
            VerificacionComprobantes.Service.StorageService<VerificacionComprobantes.Model.Person> service = new VerificacionComprobantes.Service.StorageService<VerificacionComprobantes.Model.Person>();

            VerificacionComprobantes.Model.Person person = service.Query(LiteDB.Query.EQ("Name", ((ListBoxItem)foundPerson.SelectedItem).Content.ToString() ), "personas").First();

            if (person == null)
            {
                return;
            }

            System.Globalization.TextInfo textInfo = new System.Globalization.CultureInfo("es-AR", false).TextInfo;

            var html = new HtmlAgilityPack.HtmlDocument();
            html.Load(@templateChooser.Text);

            var document = html.DocumentNode;
            document.QuerySelector("#nombre").InnerHtml = textInfo.ToTitleCase(person.Name);

            foreach (VerificacionComprobantes.Model.Operation operation in person.Operations)
            {
                if (string.IsNullOrEmpty(operation.Voucher))
                {
                    var row = HtmlAgilityPack.HtmlNode.CreateNode("<tr>");
                    var cell = HtmlAgilityPack.HtmlNode.CreateNode("<td style='border: 1px solid #ddd; padding: 8px;'>");

                    cell.InnerHtml = operation.FacturationDate;
                    row.AppendChild(cell);

                    cell = HtmlAgilityPack.HtmlNode.CreateNode("<td style='border: 1px solid #ddd; padding: 8px;'>");

                    cell.InnerHtml = operation.Total;
                    row.AppendChild(cell);

                    cell = HtmlAgilityPack.HtmlNode.CreateNode("<td style='border: 1px solid #ddd; padding: 8px;'>");

                    cell.InnerHtml = operation.Entity;

                    row.AppendChild(cell);

                    document.QuerySelector("#pendientes").AppendChild(row);
                }
            }

            document.QuerySelector("#firmante").InnerHtml = remitente.Text;

            VerificacionComprobantes.Service.EmailServer server = new VerificacionComprobantes.Service.EmailServer();

            server.Send("pedro.zuppelli@gmail.com", ccEmails.Text,
                config.Email.Header, html.DocumentNode.OuterHtml);
        }

        private void updateEmail(object sender, RoutedEventArgs e)
        {
            if (null == foundPerson.SelectedItem) return;

            VerificacionComprobantes.Service.StorageService<VerificacionComprobantes.Model.Person> service = new VerificacionComprobantes.Service.StorageService<VerificacionComprobantes.Model.Person>();

            VerificacionComprobantes.Model.Person person = service.Query(LiteDB.Query.EQ("Name", ((ListBoxItem)foundPerson.SelectedItem).Content.ToString()), "personas").First();
            
            if (string.IsNullOrEmpty(person.Email) || !person.Email.Equals(userEmail.Text))
            {
                person.Email = userEmail.Text;
                service.Update(person);
                SelectPerson(sender, null);
            }
        }

        private void updateEmailKey(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
                updateEmail(sender, null);
        }

        private void saveConfig(object sender, RoutedEventArgs e)
        {
            VerificacionComprobantes.Service.ConfigurationServer configServer = VerificacionComprobantes.Service.ConfigurationServer.Instance;
            VerificacionComprobantes.Service.ConfigurationServer.Configuration config = new VerificacionComprobantes.Service.ConfigurationServer.Configuration();

            config.Email.Server = smtp.Text;
            config.Email.Port = port.Text;
            config.Email.User = user.Text;
            config.Email.Password = pass.Password;
            config.Email.Header = header.Text;
            config.Email.Template = templateChooser.Text;
            config.Email.CC = ccEmails.Text;
            config.Email.Sender = remitente.Text;

            XLSDataSource.ColumnMappings.SHEETS = config.Sheet.Sheets = solapas.Text;
            XLSDataSource.ColumnMappings.FECHA = config.Sheet.Date = fecha.Text;
            XLSDataSource.ColumnMappings.OPERATION = config.Sheet.Operation = noperacion.Text;
            XLSDataSource.ColumnMappings.VOUCHER = config.Sheet.Voucher = voucher.Text;
            XLSDataSource.ColumnMappings.NAME = config.Sheet.Name = nombre.Text;
            XLSDataSource.ColumnMappings.TOTAL = config.Sheet.Total = total.Text;

            configServer.Store(config);
        }
    }
}
