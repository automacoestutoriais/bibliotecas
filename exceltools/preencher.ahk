#SingleInstance Force
#Include <ExcelTools>

Esc::

abrir("C:\Users\denil\Desktop\base.xlsx", 0)

Loop
{
	if ((celulaAtual := capturar("A" . (A_Index + 1))) != "")
	{
		Clipboard := celulaAtual
		Send, ^v
		Sleep, 2500
		Send, {Tab 7}
		Sleep, 500
		Send, {Enter}
		Sleep, 500
		Send, +{Tab 7}
		Sleep, 500
		Send, ^a
		Sleep, 500
		escrever("B" . (A_Index + 1), Clipboard)
		salvar()
	}
	else
	{
		break
	}
}

sair()

MsgBox, 64, Sucesso, Processo finalizado!

return
	