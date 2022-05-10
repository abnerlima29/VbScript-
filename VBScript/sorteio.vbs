dim n,i,nome(6) 'Qtde de posições do vetor
call carregar_nomes
sub carregar_nomes()
nome(1)="Moquidésia"
nome(2)="Jurema"
nome(3)="Lindolfo"
nome(4)="Astolfo"
nome(5)="Kleber"
nome(6)="Tramontinaldo"
i=1
do while i<=10
'for i=6 to 1 step -1 'Estrutura de Repetição
    randomize(second(time))
	n=int(rnd * 6) + 1
    msgbox(nome(n)),vbinformation+vbOKOnly,"Qtde Sorteios: "& i &""
'next 
i=i+1
loop
msgbox("Fim da Leitura!"),vbInformation+vbOKOnly,"AVISO"
end sub
