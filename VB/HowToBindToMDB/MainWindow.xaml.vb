Imports DevExpress.Xpf.PivotGrid
Imports System.Data
Imports System.Data.OleDb
Imports System.Windows

Namespace HowToBindToMDB

    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Public Partial Class MainWindow
        Inherits Window

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Create a connection object.
            Dim connection As OleDbConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NWIND.MDB")
            ' Create a data adapter.
            Dim adapter As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM SalesPerson", connection)
            ' Create and fill a dataset.
            Dim sourceDataSet As DataSet = New DataSet()
            adapter.Fill(sourceDataSet, "SalesPerson")
            ' Assign the data source to the PivotGrid control.
            pivotGridControl1.DataSource = sourceDataSet.Tables("SalesPerson")
            pivotGridControl1.DataProcessingEngine = DataProcessingEngine.Optimized
            pivotGridControl1.BeginUpdate()
            AddField("Country", FieldArea.RowArea, "Country", 0)
            AddField("Person", FieldArea.RowArea, "Sales Person", 1)
            AddField("Category", FieldArea.ColumnArea, "CategoryName", 0)
            AddField("Year", FieldArea.ColumnArea, "OrderDate", 1)
            AddField("Price", FieldArea.DataArea, "Extended Price", 0)
            Dim orderDateBinding As DataBinding = pivotGridControl1.Fields("OrderDate").DataBinding
            TryCast(orderDateBinding, DataSourceColumnBinding).GroupInterval = FieldGroupInterval.DateYear
            pivotGridControl1.EndUpdate()
        End Sub

        Private Sub AddField(ByVal caption As String, ByVal area As FieldArea, ByVal columnName As String, ByVal index As Integer)
            Dim field As PivotGridField = pivotGridControl1.Fields.Add()
            field.Caption = caption
            field.Area = area
            field.DataBinding = New DataSourceColumnBinding(columnName)
            field.AreaIndex = index
        End Sub
    End Class
End Namespace
