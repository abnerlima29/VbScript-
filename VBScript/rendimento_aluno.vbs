Dim n1,n2,n3,media,situacao 'Declaração de Variáveis
Dim resp,audio

call carregar_voz
sub carregar_voz()
set audio=CreateObject("SAPI.SPVOICE")
audio.volume=100
audio.rate = 2 'Velocidade da voz
call entrada_notas
end Sub

sub entrada_notas()

'Entrada de Dados
n1=cdbl(inputbox("Digite a nota 01","AVISO"))
n2=cdbl(inputbox("Digite a nota 02","AVISO"))
n3=cdbl(inputbox("Digite a nota 03","AVISO"))

'Processamento
media=round((n1+n2+n3)/3,1)
if media < 4 Then
   situacao="Reprovado"
elseif media >=4 and media < 7 then 
   situacao="Exame"
else 
   situacao="Aprovado"
end if

'Saída de Dados

'Por voz
audio.speak ("Rendimento do aluno" + vbnewline &_
             "Média Final "& media &"" + vbnewline &_
			 "Situação do aluno: "& situacao &"")
			 
'Por mensagem
resp=msgbox("Rendimento do Aluno" + vbnewline &_
            "Média Final: "& media &"" + vbnewline &_
			"Situação Aluno: "& situacao &"" + vbnewline &_
			"Novo Cálculo?",vbquestion+vbyesno,"AVISO")
if resp=vbyes Then
   call entrada_notas
Else
   wscript.quit
end if
end sub



