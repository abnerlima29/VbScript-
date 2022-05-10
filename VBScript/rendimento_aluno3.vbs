'Declaração de variáveis'todas as váriaveis assumem alfanumericas
'todas as váriaveis tem que ser criadas no top do código
'tepois que eu converto o tipo
'sempre usaremos o Dim para declarar as variaveis no inicio
'dica, sempre declarar no maximo 5 variaveis por linha
'caso queira declarar mais, usar um novo dim na linha de baixo
'call chama uma futura função
'sub - rotina que eu quero trabalhar
'inputbox - entrada de dados
'cdbl - converte a variavel em decimal
'round - com ele eu consigo definir quantas casas decimais que eu quero
'entrada
'processo
'saida
'vbnewline &_ usado para pular linha
' ("") - para concatenar um variavel usar aspas duplas


Dim n1, n2, n3, situacao, resp 'Declaração de variáveis
Dim audio

call entrada_notas
sub entrada_notas()

call carregar_voz
sub carregar_voz()
set audio=CreateObject("SAPI.SPVOICE")
audio.volume = 100 'volume da voz
audio.rate = 2 'valocidade da voz
end sub
'entrada de dados

n1=cdbl (inputbox("Digite a nota 01", AVISO"))
n2=cdbl (inputbox("Digite a nota 02", AVISO"))
n3=cdbl (inputbox("Digite a nota 03", AVISO"))

'processamento

media=round((n1+n2+n3)/3,1)
if media < 4 then
	situacao="REPROVADO"
elseif meida >=4 and media < 7 then
	situacao="Exame"
else
	situacao="Aprovado"

'Sáida de dados

audio.speak ("rendimento do aluno" + vbNewLine &_
			 "Média final "& media &"" +vbNewLine &_
			 "Situação do aluno: "& situacao &"")

resp-msdbox("Rendimento do aluno" + vbnewline &_
		"Média Final: "& media &"" + vbnewline &_
		"situação do aluno: "& situacao &"" + vbnewline &_
		"Novo cálculo?", vbquestion+vbyesno, "AVISO")
if resp=vbyes then
   call entrada_notas
else
   wscript.quit
end if
end sub