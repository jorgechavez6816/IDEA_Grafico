Sub Main
	Call ChartData()	'Resumen01.IMD
End Sub


' Datos: Gráfico de datos
Function ChartData
	Set db = Client.OpenDatabase("Resumen01.IMD")
	Set task = db.ChartData
	task.ChartTitle = "Resumen_Ventas_Producto"
	task.XFieldTitle = "Vendedores"
	task.YFieldTitle = "Total Ventas"
	task.SnapShot = TRUE
	task.NoOfSeries = 1
	task.ChartType = 1
	task.Show3DChart = TRUE
	task.ShowGrids = FALSE
	task.Legend = FALSE
	task.Criteria = ""
	task.NumOfRecords = 419
	task.XFieldName = "NUM_VENDEDOR"
	task.AddYFieldName "TOTAL_SUMA"
	task.ResultName = db.UniqueResultName("Gráfico_01")
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function