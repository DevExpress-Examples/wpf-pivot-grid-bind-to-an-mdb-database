using DevExpress.Xpf.PivotGrid;
using System.Data;
using System.Data.OleDb;
using System.Windows;

namespace HowToBindToMDB {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {
        public MainWindow() {
            InitializeComponent();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e) {
            // Create a connection object.
            OleDbConnection connection =
                new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NWIND.MDB");
            // Create a data adapter.
            OleDbDataAdapter adapter =
                new OleDbDataAdapter("SELECT * FROM SalesPerson", connection);

            // Create and fill a dataset.
            DataSet sourceDataSet = new DataSet();
            adapter.Fill(sourceDataSet, "SalesPerson");

            // Assign the data source to the PivotGrid control.
            pivotGridControl1.DataSource = sourceDataSet.Tables["SalesPerson"];


            pivotGridControl1.BeginUpdate();

            AddField("Country", FieldArea.RowArea, "Country", 0);
            AddField("Person", FieldArea.RowArea, "Sales Person", 1);
            AddField("Category", FieldArea.ColumnArea, "CategoryName", 0);
            AddField("Year", FieldArea.ColumnArea, "OrderDate", 1);
            AddField("Price", FieldArea.DataArea, "Extended Price", 0);

            DataBinding orderDateBinding = pivotGridControl1.Fields["OrderDate"].DataBinding;
            (orderDateBinding as DataSourceColumnBinding).GroupInterval = FieldGroupInterval.DateYear;

            pivotGridControl1.EndUpdate();

        }
        private void AddField(string caption, FieldArea area, string columnName, int index) {
            PivotGridField field = pivotGridControl1.Fields.Add();
            field.Caption = caption;
            field.Area = area;
            field.DataBinding = new DataSourceColumnBinding(columnName);
            field.AreaIndex = index;
        }

    }
}
