dim farol,cor,Resp
call inicio
sub inicio()
farol=cint(inputbox("[1] Verde" + vbnewline &_
      "[2] Amarelo" + vbnewline &_
	  "[3] Vermelho" + vbNewLine &_
	  "[0 ou 10] Encerrar Script","CORES DO SEM�FORO"))
select case farol
       case 1:
	        cor="Verde - Siga em Frente"
	   case 2:
	        cor="Amarelo - Aten��o"
       case 3:
	        cor="Vermelho - Pare"
	   case 0,10:
	        resp=msgbox("Deseja Encerrar?",vbquestion+vbyesno,"ATEN��O")
			if resp=vbyes Then
			   wscript.quit
			Else
			   call inicio
			end if
	   case Else
	        cor="Op��o Inv�lida!"
end Select
msgbox(""& cor &""),vbInformation+vbOKOnly,"CORES DO SEM�FORO"
call inicio
end sub








	        