Imports System.Data.SqlClient
Imports System.Data.SqlDbType
Imports System.Data.Common
Imports System.Configuration

Public Class Form1
    Dim usertable As New DataTable
    Dim sda As New SqlDataAdapter
    Dim sqlconn As SqlConnection = getdatabaseconnection("AdventureWorks2012")

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Using sqlconn = getdatabaseconnection("AdventureWorks2012")
            sqlconn.Open()
            sda.SelectCommand = New SqlCommand("select * from person.person", sqlconn)
            sda.Fill(usertable)
            BindingSource1.DataSource = usertable
            DataGridView1.DataSource = BindingSource1
            BindingNavigator1.BindingSource = BindingSource1
            For Each a As DataColumn In usertable.Columns
                ListBox1.Items.Add(a.ColumnName)
            Next
            usertable.PrimaryKey = New DataColumn() {usertable.Columns("BusinessEntityID")}


        End Using
    End Sub

    Public Function getdatabaseconnection(name As String) As DbConnection
        Dim settings As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(name)
        Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(settings.ProviderName)
        Dim conn As DbConnection = factory.CreateConnection
        conn.ConnectionString = settings.ConnectionString
        Return conn
    End Function

    Private Sub BindingNavigatorDeleteItem_Click(sender As Object, e As EventArgs) Handles BindingNavigatorDeleteItem.Click
        Dim currentindex As Integer = BindingSource1.Position
        Dim myrow As DataRow = usertable.Rows.Find(usertable.Rows(currentindex - 1).Item("BusinessEntityID"))
        myrow.Delete()
        Using sqlconn = getdatabaseconnection("AdventureWorks2012")
            sqlconn.Open()
            Dim deletecommand As SqlCommand = New SqlCommand("delete from person.person where BusinessEntityID = @BusinessEntityID")
            deletecommand.Connection = sqlconn
            deletecommand.Parameters.Add("@BusinessEntityID", SqlDbType.Int, 10, "BusinessEntityID")

            sda.DeleteCommand = deletecommand
            Try
                sda.Update(usertable)
            Catch ex As SqlException
                MsgBox(ex.Message)

            End Try

        End Using







    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit

        Dim currentindex As Integer = BindingSource1.Position
        Dim myrow As DataRow = usertable.Rows.Find(usertable.Rows(currentindex).Item("BusinessEntityID"))
        MsgBox(myrow(e.ColumnIndex))


        Using sqlconn = getdatabaseconnection("AdventureWorks2012")
            sqlconn.Open()
            Dim updatecommand As SqlCommand = New SqlCommand("update person.person set " + usertable.Columns(e.ColumnIndex).ColumnName + "=" + "'" + myrow(e.ColumnIndex) + "'" + " where BusinessEntityID = @BusinessEntityID")
            updatecommand.Connection = sqlconn
            updatecommand.Parameters.Add("@BusinessEntityID", SqlDbType.Int, 10, "BusinessEntityID")

            sda.updatecommand = updatecommand
            Try
                sda.Update(usertable)
                usertable.AcceptChanges()
            Catch ex As SqlException
                MsgBox(ex.Message)

            End Try

        End Using
    End Sub

    Private Sub DataGridView1_RowStateChanged(sender As Object, e As DataGridViewRowStateChangedEventArgs) Handles DataGridView1.RowStateChanged

    End Sub

    Private Sub DataGridView1_UserAddedRow(sender As Object, e As DataGridViewRowEventArgs) Handles DataGridView1.UserAddedRow

    End Sub
End Class
