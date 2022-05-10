dim farol,cor,Resp
call inicio
sub inicio()
farol=cint(inputbox("[1] Verde" + vbnewline &_
      "[2] Amarelo" + vbnewline &_
	  "[3] Vermelho" + vbNewLine &_
	  "[0 ou 10] Encerrar Script","CORES DO SEMÁFORO"))
select case farol
       case 1:
	        cor="Verde - Siga em Frente"
	   case 2:
	        cor="Amarelo - Atenção"
       case 3:
	        cor="Vermelho - Pare"
	   case 0,10:
	        resp=msgbox("Deseja Encerrar?",vbquestion+vbyesno,"ATENÇÃO")
			if resp=vbyes Then
			   wscript.quit
			Else
			   call inicio
			end if
	   case Else
	        cor="Opção Inválida!"
end Select
msgbox(""& cor &""),vbInformation+vbOKOnly,"CORES DO SEMÁFORO"
call inicio
end sub








	        