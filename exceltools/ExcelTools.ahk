abrir(caminho, visibilidade:=1) {
	global eab := ComObjCreate("Excel.Application")
	global ab:= eab.Workbooks.Open(caminho)
	eab.Visible := visibilidade = 1 ? True : False
}

capturar(celula) {
	global eab, eco, ecr
	return eab ? eab.Range(celula).Value : eco ? eco.Range(celula).Value : ecr.Range(celula).Value
}

conectar() {
	global eco := ComObjActive("Excel.Application")
	global co := eco.ActiveWorkbook
}

criar(visibilidade:=1) {
	global ecr := ComObjCreate("Excel.Application")
	global cr := ecr.Workbooks.Add
	ecr.Visible := visibilidade = 1 ? True : False
}

escrever(celula, valor) {
	e := ComObjActive("Excel.Application")
	e.Range(celula).Value := valor
}

sair() {
	global eab, eco, ecr
	eab ? eab.Quit() : eco ? eco.Quit() : ecr.Quit()
}

salvar(caminho:=false) {
	global ab, cr, co
	ab ? caminho ? ab.SaveAs(caminho) : ab.Save() : cr ? caminho ? cr.SaveAs(caminho) : cr.Save() : co ? caminho ? co.SaveAs(caminho) : co.Save()
}

OnError("erro")
erro(e) {
	MsgBox, 48, Aviso, % "Erro: " . e.Message . "`n`nScript: " . e.File . "`n`nComando: " e.What . "`n`nLinha: " . e.Line
	return true
}