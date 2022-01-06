SELECT ID_NatOper
	,Fil_NatOper
	,CFOP_NatOper
	,Descr_NatOper
	,STICMS_NatOper
	,STIPI_NatOper
	,STPC_NatOper
-- SELECT COUNT(*)	-- SELECT * -- DELETE
FROM tblNatOp;

SET IDENTITY_INSERT tblNatOp ON

INSERT INTO tblNatOp (
	ID_NatOper
	,Fil_NatOper
	,CFOP_NatOper
	,Descr_NatOper
	,STICMS_NatOper
	,STIPI_NatOper
	,STPC_NatOper
	)
SELECT *
FROM (
	VALUES (
		1
		,'PSP'
		,'5.102'
		,'Revenda'
		,'00'
		,'50'
		,'01'
		)
		,(
		2
		,'PSP'
		,'6.108'
		,'Revenda p/ Consumo (PF)'
		,'00'
		,'50'
		,'01'
		)
		,(
		3
		,'PSP'
		,'5.933'
		,'Presta��o de Servi�os'
		,'41'
		,'53'
		,'01'
		)
		,(
		4
		,'PSP'
		,'6.933'
		,'Presta��o de Servi�os'
		,'41'
		,'53'
		,'01'
		)
		,(
		5
		,'PSP'
		,'1.102'
		,'Compra p/ Revenda'
		,'00'
		,'00'
		,'08'
		)
		,(
		7
		,'PSP'
		,'5.202'
		,'Devolu��o de Compra'
		,''
		,''
		,''
		)
		,(
		8
		,'PSP'
		,'6.202'
		,'Devolu��o de Compra'
		,''
		,''
		,''
		)
		,(
		9
		,'PSP'
		,'1.949'
		,'Entrada de Produtod p/ Mostru�rio'
		,''
		,''
		,''
		)
		,(
		11
		,'PSP'
		,'1.917'
		,'Entrada de Mercadoria em Consigna��o'
		,''
		,''
		,''
		)
		,(
		13
		,'PSP'
		,'5.918'
		,'Devolu��o de Recebimento em Consigna��o'
		,''
		,''
		,''
		)
		,(
		14
		,'PSP'
		,'5.114'
		,'Venda de Consigna��o'
		,'41'
		,'53'
		,'01'
		)
		,(
		15
		,'PSP'
		,'6.102'
		,'Revenda'
		,'00'
		,'50'
		,'01'
		)
		,(
		16
		,'PSP'
		,'3.102'
		,'Importa��o p/ Revenda'
		,'00'
		,'00'
		,'01'
		)
		,(
		17
		,'PSP'
		,'6.949'
		,'Troca em Garantia'
		,'00'
		,'53'
		,'08'
		)
		,(
		18
		,'PSP'
		,'5.917'
		,'Remessa em Consigna��o'
		,'41'
		,'50'
		,'08'
		)
		,(
		19
		,'PSP'
		,'5.949'
		,'Troca em Garantia'
		,'00'
		,'53'
		,'08'
		)
		,(
		20
		,'PSP'
		,'6.917'
		,'Remessa em Consigna��o'
		,'41'
		,'50'
		,'08'
		)
		,(
		21
		,'PSP'
		,'1.949'
		,'Entrada p/ Garantia'
		,'90'
		,'03'
		,'08'
		)
		,(
		22
		,'PSP'
		,'1.202'
		,'Devolu��o de Venda'
		,'00'
		,'00'
		,'99'
		)
		,(
		23
		,'PSP'
		,'2.202'
		,'Devolu��o de Venda'
		,'00'
		,'00'
		,'99'
		)
		,(
		24
		,'PSP'
		,'6.910'
		,'Remessa de Doa��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		25
		,'PSP'
		,'5.910'
		,'Remessa de Doa��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		26
		,'PSP'
		,'6.918'
		,'Devolu��o de Recebimento em Consigna��o'
		,''
		,''
		,''
		)
		,(
		28
		,'PSP'
		,'5.949'
		,'Remessa p/ Mostru�rio'
		,''
		,''
		,''
		)
		,(
		29
		,'PSP'
		,'1.915'
		,'Entrada p/ Conserto ou Reparo'
		,'40'
		,'03'
		,'08'
		)
		,(
		30
		,'PSP'
		,'2.915'
		,'Entrada p/ Conserto ou Reparo'
		,'50'
		,'05'
		,'98'
		)
		,(
		31
		,'PSP'
		,'5.905'
		,'Rem. p/ Armazem'
		,'41'
		,'53'
		,'08'
		)
		,(
		32
		,'PSP'
		,'6.114'
		,'Venda de Consigna��o'
		,'41'
		,'53'
		,'01'
		)
		,(
		33
		,'PSP'
		,'5.102'
		,'Complemento de ICMS'
		,''
		,''
		,''
		)
		,(
		34
		,'PSP'
		,'6.949'
		,'Remessa p/ Reposi��o'
		,''
		,''
		,''
		)
		,(
		35
		,'PSP'
		,'5.949'
		,'Remessa p/ Reposi��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		36
		,'PSP'
		,'5.908'
		,'Remessa em Comodato'
		,'40'
		,'55'
		,'07'
		)
		,(
		37
		,'PSP'
		,'6.908'
		,'Remessa em Comodato'
		,''
		,''
		,''
		)
		,(
		39
		,'PSP'
		,'6.916'
		,'Retorno Mercadoria Recebida p/ Conserto ou Reparo'
		,'50'
		,'55'
		,'08'
		)
		,(
		40
		,'PSP'
		,'5.916'
		,'Retorno Mercadoria Recebida p/ Conserto ou Reparo'
		,'40'
		,'53'
		,'08'
		)
		,(
		41
		,'PSP'
		,'6.102'
		,'Complemento de ICMS'
		,''
		,''
		,''
		)
		,(
		42
		,'PSP'
		,'6.108'
		,'Complemento de ICMS'
		,''
		,''
		,''
		)
		,(
		43
		,'PSP'
		,'2.949'
		,'Entrada p/ Garantia'
		,'90'
		,'03'
		,'08'
		)
		,(
		44
		,'PSP'
		,'5.949'
		,'Baixa de Estoque'
		,''
		,''
		,''
		)
		,(
		45
		,'PSP'
		,'1.909'
		,'Retorno de Remessa por Contrato Comodato'
		,''
		,''
		,''
		)
		,(
		46
		,'PSP'
		,'2.909'
		,'Retorno de Remessa por Contrato Comodato'
		,''
		,''
		,''
		)
		,(
		47
		,'PSP'
		,'5.914'
		,'Remessa p/ Exposi��o ou Feira'
		,'40'
		,'55'
		,'08'
		)
		,(
		48
		,'PSP'
		,'6.914'
		,'Remessa p/ Exposi��o ou Feira'
		,'00'
		,'53'
		,'08'
		)
		,(
		49
		,'PSP'
		,'1.949'
		,'Outras Entradas'
		,'00'
		,'50'
		,'08'
		)
		,(
		50
		,'PSP'
		,'5.949'
		,'Outras Sa�das'
		,'00'
		,'50'
		,'08'
		)
		,(
		51
		,'PSP'
		,'5.912'
		,'Remessa p/ demonstra��o'
		,'50'
		,'50'
		,'08'
		)
		,(
		52
		,'PSP'
		,'6.912'
		,'Remessa p/ demonstra��o'
		,'50'
		,'50'
		,'08'
		)
		,(
		53
		,'PSP'
		,'5.913'
		,'Retorno de Recebimento p/ Demonstra��o'
		,'41'
		,'99'
		,'08'
		)
		,(
		54
		,'PSP'
		,'6.913'
		,'Retorno de Recebimento p/ Demonstra��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		55
		,'PSP'
		,'5.949'
		,'Remessa p/ Reposi��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		56
		,'PSP'
		,'1.916'
		,'Retorno de Conserto ou Reparo'
		,''
		,''
		,''
		)
		,(
		57
		,'PSP'
		,'1.949'
		,'Retorno de Remessa p/ Mostru�rio'
		,''
		,''
		,''
		)
		,(
		58
		,'PSP'
		,'1.918'
		,'Devolu��o de Remessa em Consigna��o'
		,''
		,''
		,''
		)
		,(
		59
		,'PSP'
		,'2.918'
		,'Devolu��o de Remessa em Consigna��o'
		,'00'
		,'00'
		,'08'
		)
		,(
		60
		,'PSP'
		,'5.403'
		,'Venda Adquirida de Terceiros c/ ST - (Substituto)'
		,'10'
		,'50'
		,'01'
		)
		,(
		63
		,'PSP'
		,'1.102*1.403'
		,'Compra/ Compra c/ ST'
		,''
		,''
		,''
		)
		,(
		64
		,'PSP'
		,'5.102'
		,'Revenda p/ Consumo (PJ)'
		,'00'
		,'50'
		,'01'
		)
		,(
		65
		,'PSP'
		,'6.102'
		,'Revenda p/ Consumo (PJ)'
		,'00'
		,'50'
		,'01'
		)
		,(
		66
		,'PSP'
		,'1.913'
		,'Retorno de remessa p/ demonstra��o'
		,'50'
		,'00'
		,'08'
		)
		,(
		67
		,'PSP'
		,'1.411'
		,'Devolu��o de Venda c/ ST'
		,'10'
		,'00'
		,'08'
		)
		,(
		68
		,'PSP'
		,'5.949'
		,'- Inativo -Remessa para empr�stimo'
		,'00'
		,'50'
		,'08'
		)
		,(
		69
		,'PSP'
		,'6.403'
		,'Venda Adquirida de Terceiros c/ ST - (Substituto)'
		,'10'
		,'50'
		,'01'
		)
		,(
		70
		,'PSP'
		,'2.411*2.202'
		,'Devolu��o de Venda'
		,''
		,''
		,''
		)
		,(
		71
		,'PSP'
		,'3.949'
		,'Inativo Importa��o p/ Garantia - Pagamento Invoice'
		,'00'
		,'00'
		,'01'
		)
		,(
		72
		,'PSP'
		,'5.910'
		,'Remessa p/ Distribui��o/ Brinde'
		,''
		,''
		,''
		)
		,(
		73
		,'PSP'
		,'3.556'
		,'Importa��o p/ Uso ou Consumo'
		,'41'
		,'03'
		,'08'
		)
		,(
		74
		,'PSP'
		,'6.152'
		,'Transferencia de Mercadorias'
		,'00'
		,'50'
		,'08'
		)
		,(
		75
		,'PSP'
		,'6.949'
		,'Retorno em Garantia'
		,'90'
		,'99'
		,'99'
		)
		,(
		76
		,'PSP'
		,'1.409'
		,'Entrada Transfer�ncia c/ ST'
		,''
		,''
		,''
		)
		,(
		77
		,'PSP'
		,'2.409'
		,'Entrada deTransfer�ncia c/ ST'
		,''
		,''
		,''
		)
		,(
		78
		,'PSP'
		,'3.949'
		,'Exposi��o'
		,''
		,''
		,''
		)
		,(
		79
		,'PSP'
		,'2.913'
		,'Retorno de remessa p/ demonstra��o'
		,''
		,''
		,''
		)
		,(
		80
		,'PSP'
		,'1.906'
		,'Ret Mercadoria Rem p/ Dep�sito Fech ou Armz Geral'
		,'41'
		,'03'
		,'07'
		)
		,(
		81
		,'PSP'
		,'1.353'
		,'Transporte Rodovi�rio'
		,'00'
		,'03'
		,'07'
		)
		,(
		82
		,'PSP'
		,'2.949'
		,'Outras Entradas'
		,'00'
		,'00'
		,'08'
		)
		,(
		83
		,'PSP'
		,'1.910'
		,'Entrada de Bonifica��o, Doa��o ou Brinde'
		,''
		,''
		,''
		)
		,(
		84
		,'PSP'
		,'1.000'
		,'Entrada ou Aquisi��o de Servi�o'
		,'41'
		,'03'
		,'07'
		)
		,(
		85
		,'PSP'
		,'1.556'
		,'Compra p/ Uso ou Consumo'
		,'90'
		,'49'
		,'07'
		)
		,(
		86
		,'PSP'
		,'2.000'
		,'Entrada ou Aquisi��o de Servi�o'
		,'41'
		,'03'
		,'07'
		)
		,(
		87
		,'PSP'
		,'1.303'
		,'Aquisi��o de Servi�o de Comunica��o'
		,'00'
		,'03'
		,'08'
		)
		,(
		89
		,'PSP'
		,'2.556'
		,'Compra p/ Uso ou Consumo'
		,'90'
		,'49'
		,'07'
		)
		,(
		90
		,'PSP'
		,'5.949'
		,'Emprestimo de Mercadoria'
		,'00'
		,'53'
		,'08'
		)
		,(
		91
		,'PSP'
		,'2.152'
		,'Entrada Transfer�ncia p/ Comercializa��o'
		,'00'
		,'00'
		,'07'
		)
		,(
		92
		,'PSP'
		,'2.914'
		,'Retorno de Remessa p/ Exposi��o/Feira'
		,'00'
		,'03'
		,'08'
		)
		,(
		93
		,'PSP'
		,'2.910'
		,'Entrada de Bonifica��o, Doa��o ou Brinde'
		,'00'
		,'00'
		,'08'
		)
		,(
		94
		,'PSP'
		,'6.949'
		,'Complemento'
		,'00'
		,'50'
		,'08'
		)
		,(
		95
		,'PSP'
		,'2.353'
		,'Transporte Rodovi�rio'
		,'00'
		,'03'
		,'07'
		)
		,(
		96
		,'PSP'
		,'5.413'
		,'Devolu��o de Compra p/ Uso ou Consumo c/ ST'
		,'60'
		,'53'
		,'08'
		)
		,(
		97
		,'PSP'
		,'7.949'
		,'Remessa p/ An�lise'
		,'41'
		,'53'
		,'49'
		)
		,(
		98
		,'PSP'
		,'3.949'
		,'Reposi��o'
		,''
		,''
		,''
		)
		,(
		99
		,'PSP'
		,'2.912'
		,' Entrada recebida p/ demonstra��o'
		,'00'
		,'00'
		,'08'
		)
		,(
		100
		,'PSP'
		,'2.102'
		,'Compra p/ Revenda'
		,'00'
		,'00'
		,'08'
		)
		,(
		101
		,'PSP'
		,'5.915'
		,'Remessa p/ Conserto ou Reparo'
		,'41'
		,'53'
		,'08'
		)
		,(
		102
		,'PSP'
		,'1.912'
		,'Entrada recebida p/ demonstra��o'
		,'41'
		,'49'
		,'08'
		)
		,(
		103
		,'PSP'
		,'6.949'
		,'Emprestimo de Mercadoria'
		,'00'
		,'53'
		,'08'
		)
		,(
		104
		,'PSP'
		,'1.908'
		,'Entrada de Bem por Contrato Comodato'
		,'41'
		,'03'
		,'08'
		)
		,(
		105
		,'PSP'
		,'2.411'
		,'Devolu��o de Venda c/ ST'
		,'10'
		,'00'
		,'08'
		)
		,(
		106
		,'PSP'
		,'1.914'
		,'Retorno de Remessa p/ Exposi��o/Feira'
		,'40'
		,'05'
		,'08'
		)
		,(
		107
		,'PSP'
		,'6.909'
		,'Retorno de Recebimento em Comodato'
		,'41'
		,'53'
		,'08'
		)
		,(
		108
		,'PSP'
		,'1.949'
		,'Outras Entradas'
		,'00'
		,'00'
		,'08'
		)
		,(
		110
		,'PSP'
		,'7.102'
		,'Exporta��o'
		,'41'
		,'53'
		,'08'
		)
		,(
		111
		,'PSP'
		,'6.557'
		,'Transf. p/ Uso e Consumo'
		,'00'
		,'50'
		,'08'
		)
		,(
		112
		,'PSP'
		,'5.557'
		,'Transferencia de Material p/ Uso e Consumo'
		,'00'
		,'50'
		,'08'
		)
		,(
		113
		,'PSP'
		,'6.110'
		,'Venda Adq. Terc. SUFRAMA (IPI)'
		,'00'
		,'55'
		,'01'
		)
		,(
		115
		,'PSP'
		,'2.557'
		,'Entrada de Transferencia p/ Uso ou Consumo'
		,'00'
		,'00'
		,'08'
		)
		,(
		116
		,'PSP'
		,'1.910'
		,'Devolu��o de Doa��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		117
		,'PSP'
		,'5.405'
		,'Venda Adquirida de Terceiros c/ ST - (Substitu�do)'
		,'60'
		,'53'
		,'01'
		)
		,(
		118
		,'PSP'
		,'6.404'
		,'Venda Adquirida de Terceiros c/ ST'
		,'60'
		,'50'
		,'01'
		)
		,(
		119
		,'PSP'
		,'5.949'
		,'Regulariza��o Mudan�a de Endere�o'
		,'41'
		,'53'
		,'08'
		)
		,(
		120
		,'PSP'
		,'1.557'
		,'Devolu��o de Transf. de Material p/ Uso ou Consumo'
		,'00'
		,'00'
		,'08'
		)
		,(
		121
		,'PSP'
		,'5.152'
		,'Transfer�ncia de Mercadorias'
		,'00'
		,'50'
		,'08'
		)
		,(
		122
		,'PSP'
		,'2.152'
		,'Complemento de Valores - Transfer�ncia'
		,'00'
		,'00'
		,'08'
		)
		,(
		123
		,'PSP'
		,'1.551'
		,'Compra de Bem p/ Ativo Imobilizado'
		,'90'
		,'49'
		,'07'
		)
		,(
		124
		,'PSP'
		,'2.949'
		,'Devolu��o Troca de Garantia'
		,'00'
		,'00'
		,'08'
		)
		,(
		125
		,'PSP'
		,'3.910'
		,'Entrada de bonifica��o, doa��o ou brinde'
		,'00'
		,'03'
		,'08'
		)
		,(
		126
		,'PSP'
		,'1.949'
		,'Retorno de Emprestimo'
		,'40'
		,'02'
		,'08'
		)
		,(
		127
		,'PSP'
		,'1.403'
		,'Compra c/ ST'
		,'60'
		,'03'
		,'70'
		)
		,(
		128
		,'PSP'
		,'2.551'
		,'Compra de Bem p/ Ativo Imobilizado'
		,'90'
		,'49'
		,'07'
		)
		,(
		129
		,'PSP'
		,'5.927'
		,'Baixa de estoque - Roubo/ Furto'
		,'41'
		,'99'
		,'49'
		)
		,(
		130
		,'PSP'
		,'3.949'
		,'Empr�stimo'
		,'41'
		,'03'
		,'08'
		)
		,(
		131
		,'PSP'
		,'7.949'
		,'Retorno de Empr�stimo'
		,'41'
		,'53'
		,'08'
		)
		,(
		132
		,'PSP'
		,'6.206'
		,'Anula��o de Valores Relativo a Serv. de Transporte'
		,'90'
		,'99'
		,'08'
		)
		,(
		133
		,'PSP'
		,'1.407'
		,'Compra de Bem p/ Uso ou Consumo c/ ST'
		,'60'
		,'49'
		,'07'
		)
		,(
		134
		,'PSP'
		,'1.406'
		,'Compra de Bem p/ Ativo Imobilizado c/ ST'
		,'90'
		,'49'
		,'07'
		)
		,(
		135
		,'PSP'
		,'5.102'
		,'Revenda (Adquirente Industrial)'
		,'00'
		,'50'
		,'01'
		)
		,(
		136
		,'PSP'
		,'6.102'
		,'Revenda (Adquirente Industrial)'
		,'00'
		,'50'
		,'01'
		)
		,(
		137
		,'PSP'
		,'3.102'
		,'Complemento de Importa��o'
		,'00'
		,'00'
		,'01'
		)
		,(
		138
		,'PSP'
		,'1.949'
		,'Devolu��o de Troca em Garantia'
		,'00'
		,'00'
		,'08'
		)
		,(
		139
		,'PSP'
		,'5.554'
		,'Remessa de Ativo p/ Uso Fora do Estabelecimento'
		,'50'
		,'53'
		,'08'
		)
		,(
		140
		,'PSP'
		,'1.554'
		,'Ret. de Ativo p/ Uso Fora do Estabelecimento'
		,'40'
		,'02'
		,'08'
		)
		,(
		141
		,'PSP'
		,'5.910'
		,'Remessa em Bonifica��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		142
		,'PSP'
		,'6.910'
		,'Remessa em Bonifica��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		143
		,'PSP'
		,'2.910'
		,'Devolu��o de Bonifica��o'
		,'00'
		,'00'
		,'74'
		)
		,(
		144
		,'PSP'
		,'1.910'
		,'Devolu��o de Bonifica��o'
		,'00'
		,'00'
		,'74'
		)
		,(
		145
		,'PSP'
		,'2.949'
		,'Retorno de Emprestimo'
		,'00'
		,'00'
		,'08'
		)
		,(
		146
		,'PSP'
		,'3.949'
		,'Simples Remessa'
		,'41'
		,'03'
		,'08'
		)
		,(
		147
		,'PSP'
		,'1.652'
		,'Compra de Combust�vel/ Lubrificante p/ Comercializ'
		,'60'
		,'03'
		,'70'
		)
		,(
		148
		,'PSP'
		,'1.653'
		,'Compra de Combust�vel/ Lubrificante p/ Uso ou Cons'
		,'40'
		,'49'
		,'07'
		)
		,(
		149
		,'PSP'
		,'1.933'
		,'Aquisi��o de Servi�o'
		,'90'
		,'49'
		,'70'
		)
		,(
		150
		,'PSP'
		,'6.110'
		,'*** N�o Usar - Venda Adq Terc SUFRAMA'
		,'00'
		,'55'
		,'06'
		)
		,(
		151
		,'PSP'
		,'5.927'
		,'Baixa de estoque - Uso Consumo'
		,'41'
		,'99'
		,'49'
		)
		,(
		152
		,'PSP'
		,'1.253'
		,'Entrada de Energia El�trica p/ Estab. Com.'
		,'41'
		,'03'
		,'08'
		)
		,(
		153
		,'PSP'
		,'5.927'
		,'Baixa de estoque - Garantia'
		,'41'
		,'99'
		,'49'
		)
		,(
		154
		,'PSP'
		,'3.949'
		,'Importa��o.'
		,'41'
		,'03'
		,'99'
		)
		,(
		157
		,'PSP'
		,'6.102'
		,'Venda Adq. Terc. SUFRAMA'
		,'00'
		,'52'
		,'01'
		)
		,(
		158
		,'PSP'
		,'2.933'
		,'Aquisi��o de Servi�o'
		,'90'
		,'49'
		,'70'
		)
		,(
		159
		,'PSP'
		,'1.922'
		,'Entrada de Simples Faturam de Compra p/ Receb Futu'
		,'90'
		,'49'
		,'70'
		)
		,(
		160
		,'PSP'
		,'2.407'
		,'Compra de Bem p/ Uso ou Consumo c/ ST'
		,'60'
		,'49'
		,'07'
		)
		,(
		161
		,'PSP'
		,'1.923'
		,'Entrada Receb do Vend Remetente em Venda � Ordem'
		,'90'
		,'49'
		,'70'
		)
		,(
		162
		,'PSP'
		,'1.116'
		,'Compra Encomendada p/ Recebimento Futuro'
		,'90'
		,'49'
		,'70'
		)
		,(
		164
		,'PSP'
		,'6.403'
		,'Complemento de ST'
		,'10'
		,'53'
		,'49'
		)
		,(
		165
		,'PSP'
		,'5.553'
		,'Devolu��o de Compra ou Bem p/ Ativo Imobilizado'
		,''
		,''
		,''
		)
		,(
		166
		,'PSP'
		,'2.102'
		,'Compra p/ Revenda (IPI N�o Tributado)'
		,'00'
		,'03'
		,'08'
		)
		,(
		167
		,'PSP'
		,'1.102'
		,'Compra p/ Revenda (IPI N�o Tributado)'
		,'00'
		,'03'
		,'08'
		)
		,(
		168
		,'PSP'
		,'5.927'
		,'Transposi��o de Estoque - Ativo'
		,'41'
		,'99'
		,'49'
		)
		,(
		169
		,'PSP'
		,'6.949'
		,'Devolu��o Bonifica��o'
		,'00'
		,'50'
		,'99'
		)
		,(
		170
		,'PSP'
		,'3.949'
		,'Complemento de Importa��o'
		,'41'
		,'03'
		,'99'
		)
		,(
		171
		,'PSP'
		,'5.551'
		,'Venda de Bem do Ativo Imobilizado'
		,'41'
		,'53'
		,'08'
		)
		,(
		172
		,'PSP'
		,'1.933'
		,'Aquisi��o de Servi�o Tributado pelo ISSQN'
		,'90'
		,'49'
		,'70'
		)
		,(
		173
		,'PSP'
		,'5.949'
		,'Simples Remessa - Sinistro'
		,'00'
		,'50'
		,'01'
		)
		,(
		174
		,'PSP'
		,'1.403'
		,'Compra p/ Revenda c/ ST (IPI N�o Tributado)'
		,'60'
		,'03'
		,'08'
		)
		,(
		175
		,'PSP'
		,'5.556'
		,'Devolu��o de Compra de Material de Uso ou Consumo'
		,'00'
		,'53'
		,'08'
		)
		,(
		176
		,'PSP'
		,'7.202'
		,'Devolu��o de Importa��o'
		,'00'
		,'50'
		,'01'
		)
		,(
		177
		,'PSP'
		,'2.910'
		,'Entrada de Bonifica��o (IPI N�o Tributado)'
		,'00'
		,'03'
		,'08'
		)
		,(
		178
		,'PSP'
		,'1.910'
		,'Entrada Bonifica��o (IPI n�o Tributado)'
		,'00'
		,'03'
		,'08'
		)
		,(
		179
		,'PSP'
		,'5.949'
		,'Amostra p/ Teste'
		,'41'
		,'53'
		,'08'
		)
		,(
		201
		,'PES'
		,'5.102'
		,'REVENDA'
		,'00'
		,'50'
		,'01'
		)
		,(
		202
		,'PES'
		,'6.108'
		,'REVENDA p/ Consumo (PF)'
		,'00'
		,'50'
		,'01'
		)
		,(
		214
		,'PES'
		,'5.114'
		,'Venda de Consigna��o'
		,'41'
		,'53'
		,'01'
		)
		,(
		215
		,'PES'
		,'6.102'
		,'REVENDA'
		,'00'
		,'50'
		,'01'
		)
		,(
		216
		,'PES'
		,'3.102'
		,'IMPORTA��O'
		,'00'
		,'00'
		,'01'
		)
		,(
		217
		,'PES'
		,'6.949'
		,'TROCA EM GARANTIA'
		,'00'
		,'53'
		,'08'
		)
		,(
		218
		,'PES'
		,'5.917'
		,'Remessa em Consigna��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		219
		,'PES'
		,'5.949'
		,'TROCA EM GARANTIA'
		,'00'
		,'50'
		,'08'
		)
		,(
		220
		,'PES'
		,'6.917'
		,'Remessa em Consigna��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		224
		,'PES'
		,'6.910'
		,'Remessa de Doa��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		225
		,'PES'
		,'5.910'
		,'Remessa de Doa��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		228
		,'PES'
		,'5.949'
		,'Remessa p/ Mostru�rio'
		,''
		,''
		,''
		)
		,(
		231
		,'PES'
		,'5.905'
		,'Rem. p/ Armazem'
		,'41'
		,'53'
		,'08'
		)
		,(
		232
		,'PES'
		,'6.114'
		,'Venda de Consigna��o'
		,'41'
		,'53'
		,'01'
		)
		,(
		233
		,'PES'
		,'5.102'
		,'NF COMPL. DE ICMS'
		,''
		,''
		,''
		)
		,(
		235
		,'PES'
		,'5.949'
		,'Remessa p/ Reposi��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		236
		,'PES'
		,'5.908'
		,'Remessa em Comodato'
		,'40'
		,'55'
		,'07'
		)
		,(
		237
		,'PES'
		,'6.908'
		,'Remessa em Comodato'
		,''
		,''
		,''
		)
		,(
		240
		,'PES'
		,'5.916'
		,'Retorno Mercadoria Recebida Conserto/Reparo'
		,'40'
		,'53'
		,'08'
		)
		,(
		241
		,'PES'
		,'6.102'
		,'NF COMPL DE ICMS'
		,''
		,''
		,''
		)
		,(
		242
		,'PES'
		,'6.108'
		,'NF COMPL ICMS'
		,''
		,''
		,''
		)
		,(
		244
		,'PES'
		,'5.949'
		,'Baixa de Estoque'
		,''
		,''
		,''
		)
		,(
		247
		,'PES'
		,'5.914'
		,'Remessa p/ Exposi��o ou Feira'
		,'40'
		,'55'
		,'08'
		)
		,(
		248
		,'PES'
		,'6.914'
		,'Remessa p/ Exposi��o ou Feira'
		,'00'
		,'53'
		,'08'
		)
		,(
		250
		,'PES'
		,'5.949'
		,'Outras Sa�das'
		,'41'
		,'53'
		,'08'
		)
		,(
		251
		,'PES'
		,'5.912'
		,'Rem. p/ demonstra��o'
		,'50'
		,'50'
		,'08'
		)
		,(
		252
		,'PES'
		,'6.912'
		,'Rem. p/demonstra��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		253
		,'PES'
		,'5.913'
		,'Ret.merc.demonstra��o'
		,''
		,''
		,''
		)
		,(
		254
		,'PES'
		,'6.913'
		,'Ret.merc.demonstra��o'
		,''
		,''
		,''
		)
		,(
		255
		,'PES'
		,'5.949'
		,'REM.P/CONSERTO'
		,''
		,''
		,''
		)
		,(
		260
		,'PES'
		,'5.403'
		,'Venda Adquirida de Terceiros c/ ST - (Substituto)'
		,'10'
		,'50'
		,'01'
		)
		,(
		264
		,'PES'
		,'5.102'
		,'REVENDA p/ Consumo (PJ)'
		,'20'
		,'50'
		,'01'
		)
		,(
		265
		,'PES'
		,'6.102'
		,'REVENDA p/ Consumo (PJ)'
		,'00'
		,'50'
		,'01'
		)
		,(
		268
		,'PES'
		,'5.949'
		,'Emprestimo de Mercadoria'
		,'20'
		,'53'
		,'08'
		)
		,(
		269
		,'PES'
		,'6.404'
		,'Venda Adq Terc c/ Subs Trib'
		,'60'
		,'50'
		,'01'
		)
		,(
		271
		,'PES'
		,'3.949'
		,'IMPORTA��O'
		,'00'
		,'00'
		,'01'
		)
		,(
		273
		,'PES'
		,'6.152'
		,'Transf. de Mercadorias'
		,'00'
		,'50'
		,'08'
		)
		,(
		274
		,'PES'
		,'2.152'
		,'Entrada Transf. p/ Comercializa��o'
		,'00'
		,'00'
		,'07'
		)
		,(
		275
		,'PES'
		,'2.202'
		,'Devolu��o de venda'
		,'00'
		,'00'
		,'99'
		)
		,(
		276
		,'PES'
		,'1.906'
		,'Ret. mercadoria rem.p/dep�sito fech ou armz geral'
		,'41'
		,'03'
		,'07'
		)
		,(
		277
		,'PES'
		,'2.949'
		,'Entrada para Garantia'
		,'00'
		,'00'
		,'08'
		)
		,(
		278
		,'PES'
		,'1.949'
		,'Outras Entradas'
		,'90'
		,'49'
		,'08'
		)
		,(
		279
		,'PES'
		,'1.000'
		,'Entradas ou Aquisi��es de Servi�o'
		,'41'
		,'03'
		,'07'
		)
		,(
		280
		,'PES'
		,'1.300'
		,'Aquisi��es de Servi�os de Comunica��o'
		,'40'
		,'49'
		,'07'
		)
		,(
		281
		,'PES'
		,'1.353'
		,'Transporte Rodovi�rio'
		,'00'
		,'03'
		,'07'
		)
		,(
		282
		,'PES'
		,'1.556'
		,'Compra p/ Uso ou Consumo'
		,'90'
		,'49'
		,'07'
		)
		,(
		283
		,'PES'
		,'2.910'
		,'Devolu��o de Doa��o'
		,'00'
		,'00'
		,'08'
		)
		,(
		284
		,'PES'
		,'2.353'
		,'Transporte Rodovi�rio'
		,'00'
		,'03'
		,'07'
		)
		,(
		285
		,'PES'
		,'2.918'
		,'Ret. Consigna��o'
		,'00'
		,'00'
		,'08'
		)
		,(
		286
		,'PES'
		,'6.114'
		,'NOTA FISCAL COMPLEMENTAR'
		,'00'
		,'50'
		,'01'
		)
		,(
		287
		,'PES'
		,'3.102'
		,'COMPLEMENTO DE IMPORTA��O'
		,'90'
		,'49'
		,'99'
		)
		,(
		288
		,'PES'
		,'5.905'
		,'NF COMPLEMENTAR REM P/ARMZ'
		,'41'
		,'53'
		,'08'
		)
		,(
		289
		,'PES'
		,'6.102'
		,'NF COMPLEMENTAR'
		,'90'
		,'99'
		,'99'
		)
		,(
		290
		,'PES'
		,'2.556'
		,'Compra p/ Uso ou Consumo'
		,'90'
		,'49'
		,'07'
		)
		,(
		291
		,'PES'
		,'6.949'
		,'Outras Sa�das'
		,'00'
		,'50'
		,'08'
		)
		,(
		292
		,'PES'
		,'1.202'
		,'Devolu��o de venda'
		,'00'
		,'00'
		,'99'
		)
		,(
		293
		,'PES'
		,'2.912'
		,'Entrada de Demonstra��o'
		,'00'
		,'00'
		,'08'
		)
		,(
		294
		,'PES'
		,'6.152'
		,'NF Complementar (Transf. de Mercadorias)'
		,'00'
		,'50'
		,'08'
		)
		,(
		295
		,'PES'
		,'2.949'
		,'NF COMPLEMENTAR'
		,'00'
		,'00'
		,'08'
		)
		,(
		296
		,'PES'
		,'5.557'
		,'Transf.Mat.p/Consumo'
		,'00'
		,'50'
		,'08'
		)
		,(
		297
		,'PES'
		,'6.949'
		,'NF COMPLEMENTAR'
		,'00'
		,'50'
		,'08'
		)
		,(
		298
		,'PES'
		,'6.912'
		,'NF COMPLEMENTAR'
		,'00'
		,'50'
		,'08'
		)
		,(
		299
		,'PES'
		,'2.908'
		,'Entr.de Prodt em Comodato'
		,'41'
		,'03'
		,'08'
		)
		,(
		300
		,'PES'
		,'2.913'
		,'Ret. Demonstra��o'
		,'00'
		,'00'
		,'08'
		)
		,(
		301
		,'PES'
		,'6.909'
		,'Retorno de Prodt Comodato'
		,'41'
		,'53'
		,'08'
		)
		,(
		302
		,'PES'
		,'6.110'
		,'Venda Adq. Terc. SUFRAMA (IPI)'
		,'00'
		,'55'
		,'01'
		)
		,(
		304
		,'PES'
		,'6.557'
		,'Transf. p/ Uso e Consumo'
		,'00'
		,'50'
		,'08'
		)
		,(
		305
		,'PES'
		,'6.110'
		,'Venda Adq. Terc. SUFRAMA CONSUMO'
		,'00'
		,'55'
		,'06'
		)
		,(
		306
		,'PES'
		,'2.949'
		,'Devolu��o Troca de Garantia'
		,'00'
		,'00'
		,'08'
		)
		,(
		307
		,'PES'
		,'2.949'
		,'Entr.Regulariza��o Suframa'
		,'00'
		,'00'
		,'99'
		)
		,(
		308
		,'PES'
		,'7.102'
		,'REGULARIZA��O DE IMPORTA��O'
		,'00'
		,'00'
		,'01'
		)
		,(
		309
		,'PES'
		,'2.000'
		,'Entradas ou Aquisi��es de Servi�o'
		,'41'
		,'03'
		,'08'
		)
		,(
		310
		,'PES'
		,'5.102'
		,'REVENDA p/ Consumo (PJ)'
		,'20'
		,'50'
		,'01'
		)
		,(
		311
		,'PES'
		,'2.102'
		,'Compra para Revenda'
		,'00'
		,'00'
		,'01'
		)
		,(
		312
		,'PES'
		,'5.927'
		,'Baixa de estoque - Roubo/ Furto'
		,'20'
		,'50'
		,'08'
		)
		,(
		313
		,'PES'
		,'2.915'
		,'Entrada p/ conserto'
		,'41'
		,'03'
		,'08'
		)
		,(
		314
		,'PES'
		,'2.557'
		,'Entrada Transf. p/ Uso e Consumo'
		,'00'
		,'00'
		,'08'
		)
		,(
		315
		,'PES'
		,'2.914'
		,'Retorno de Remessa p/ Exposi��o/Feira'
		,'00'
		,'00'
		,'08'
		)
		,(
		316
		,'PES'
		,'6.916'
		,'Retorno Mercadoria Recebida Conserto/Reparo'
		,'40'
		,'53'
		,'08'
		)
		,(
		317
		,'PES'
		,'1.949'
		,'Entrada para Garantia'
		,'00'
		,'00'
		,'08'
		)
		,(
		318
		,'PES'
		,'6.949'
		,'Simples Remessa'
		,'41'
		,'53'
		,'08'
		)
		,(
		319
		,'PES'
		,'2.204'
		,'Devolu��o de vendas SUFRAMA'
		,'41'
		,'03'
		,'08'
		)
		,(
		320
		,'PES'
		,'6.102'
		,'REVENDA (ADQUIRENTE INDUSTRIAL)'
		,'00'
		,'50'
		,'01'
		)
		,(
		321
		,'PES'
		,'5.102'
		,'REVENDA (ADQUIRENTE INDUSTRIAL)'
		,'00'
		,'50'
		,'01'
		)
		,(
		322
		,'PES'
		,'1.949'
		,'Entrada de Importa��o por Conta e Ordem'
		,'20'
		,'00'
		,'08'
		)
		,(
		323
		,'PES'
		,'5.405'
		,'Venda Adquirida de Terceiros c/ ST - (Substitu�do)'
		,'60'
		,'50'
		,'01'
		)
		,(
		324
		,'PES'
		,'1.949'
		,'Entrada de Importa��o por Conta e Ordem (ST)'
		,'70'
		,'00'
		,'08'
		)
		,(
		325
		,'PES'
		,'6.949'
		,'TROCA EM GARANTIA SUFRAMA'
		,'40'
		,'52'
		,'08'
		)
		,(
		326
		,'PES'
		,'1.949'
		,'Entrada de Importa��o por Conta e Ordem'
		,'00'
		,'00'
		,'08'
		)
		,(
		327
		,'PES'
		,'1.102'
		,'COMPRA PARA REVENDA'
		,'00'
		,'00'
		,'08'
		)
		,(
		329
		,'PES'
		,'6.403'
		,'Venda Adq Terc c/ Subs Trib'
		,'10'
		,'50'
		,'01'
		)
		,(
		330
		,'PES'
		,'5.206'
		,'ANULACAO DE VALORES RELATIVO A SERV. DE TRANSPORTE'
		,'90'
		,'99'
		,'08'
		)
		,(
		331
		,'PES'
		,'5.910'
		,'Remessa em Bonifica��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		332
		,'PES'
		,'6.910'
		,'Remessa em Bonifica��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		333
		,'PES'
		,'1.949'
		,'Entrada Importa��o Conta Ordem Extempor�nea'
		,'20'
		,'00'
		,'08'
		)
		,(
		334
		,'PES'
		,'2.910'
		,'Devolu��o de Bonifica��o'
		,'00'
		,'00'
		,'74'
		)
		,(
		335
		,'PES'
		,'2.411'
		,'Devolu��o de venda c/ ST'
		,'10'
		,'00'
		,'50'
		)
		,(
		336
		,'PES'
		,'1.907'
		,'Ret. mercadoria rem.p/dep�sito fech ou armz geral'
		,'41'
		,'03'
		,'07'
		)
		,(
		337
		,'PES'
		,'2.933'
		,'Presta��o de Servi�o'
		,'90'
		,'49'
		,'08'
		)
		,(
		338
		,'PES'
		,'6.110'
		,'*** N�O USAR - Venda Adq Terc.SUFRAM'
		,'00'
		,'52'
		,'06'
		)
		,(
		339
		,'PES'
		,'5.926'
		,'Remessa p/ forma��o de kit'
		,'41'
		,'53'
		,'08'
		)
		,(
		340
		,'PES'
		,'1.926'
		,'Entrada  p/ forma��o de kit'
		,'41'
		,'03'
		,'08'
		)
		,(
		341
		,'PES'
		,'1.949'
		,'Outras Entradas'
		,'41'
		,'02'
		,'08'
		)
		,(
		342
		,'PES'
		,'5.949'
		,'Simples Remessa - Sinistro'
		,'00'
		,'50'
		,'01'
		)
		,(
		343
		,'PES'
		,'6.949'
		,'Simples Remessa - Sinistro'
		,'00'
		,'50'
		,'01'
		)
		,(
		344
		,'PES'
		,'6.949'
		,'Emprestimo de Mercadoria'
		,'00'
		,'53'
		,'08'
		)
		,(
		346
		,'PES'
		,'6.102'
		,'Venda Adq Terc.SUFRAMA'
		,'00'
		,'52'
		,'01'
		)
		,(
		347
		,'PES'
		,'1.411'
		,'Devolu��o de venda c/ ST'
		,'10'
		,'00'
		,'50'
		)
		,(
		348
		,'PES'
		,'1.910'
		,'Devolu��o de Bonifica��o'
		,'00'
		,'00'
		,'74'
		)
		,(
		349
		,'PES'
		,'5.556'
		,'Devolu��o de Compra de Material de Uso ou Consumo'
		,'41'
		,'53'
		,'99'
		)
		,(
		350
		,'PES'
		,'2.102'
		,'Compra para Revenda (IPI N�o Tributado)'
		,'00'
		,'03'
		,'08'
		)
		,(
		351
		,'PES'
		,'2.949'
		,'Outras Entradas'
		,'90'
		,'49'
		,'08'
		)
		,(
		352
		,'PES'
		,'2.910'
		,'Entrada bonifica��o, doa��o ou brinde'
		,''
		,''
		,''
		)
		,(
		353
		,'PES'
		,'2.910'
		,'Entrada bonific, doa��o ou brinde (IPI n�o Tribut)'
		,'00'
		,'03'
		,''
		)
		,(
		354
		,'PES'
		,'6.949'
		,'Retorno em Garantia'
		,'90'
		,'99'
		,'99'
		)
		,(
		355
		,'PES'
		,'7.102'
		,'Venda de merc. Adquirida ou recebida de terceiros'
		,'41'
		,'54'
		,'07'
		)
		,(
		356
		,'PES'
		,'1.914'
		,'Retorno de Remessa p/ Exposi��o/Feira'
		,'40'
		,'05'
		,'08'
		)
		,(
		357
		,'PES'
		,'7.202'
		,'Devolu��o de IMPORTA��O'
		,'00'
		,'50'
		,'01'
		)
		,(
		359
		,'PES'
		,'2.102'
		,'Compra para Revenda (IPI N�o Tributado) IPI BCICMS'
		,'00'
		,'03'
		,'08'
		)
		,(
		360
		,'PES'
		,'2.917'
		,'Entrada de mercadoria em consigna��o'
		,'00'
		,'03'
		,'98'
		)
		,(
		361
		,'PES'
		,'6.102'
		,'Estorno de NF-e n�o cancelada no prazo legal'
		,'00'
		,'00'
		,'99'
		)
		,(
		362
		,'PES'
		,'2.949'
		,'Estorno de NF-e n�o cancelada no prazo legal'
		,'90'
		,'49'
		,'98'
		)
		,(
		363
		,'PES'
		,'7.102'
		,'Estorno de NF-e n�o cancelada no prazo legal'
		,'00'
		,'00'
		,'01'
		)
		,(
		364
		,'PES'
		,'7.949'
		,'Cancelamento Extempor�neo'
		,'41'
		,'00'
		,'01'
		)
		,(
		365
		,'PSP'
		,'2.949'
		,'Devolu��o de Bonifica��o'
		,'00'
		,'00'
		,'74'
		)
		,(
		366
		,'PSP'
		,'1.949'
		,'Devolu��o de Bonifica��o'
		,'00'
		,'00'
		,'08'
		)
		,(
		367
		,'PES'
		,'6.918'
		,'Devolu��o de mercadoria recebida em consigna��o'
		,'00'
		,'53'
		,'99'
		)
		,(
		368
		,'PSP'
		,'5.926'
		,'Remessa p/ Desagrega��o'
		,'40'
		,'49'
		,'99'
		)
		,(
		369
		,'PSP'
		,'1.926'
		,'Entrada Decorrente de Desagrega��o'
		,'40'
		,'49'
		,'99'
		)
		,(
		370
		,'PES'
		,'6.949'
		,'Estorno de NF-e n�o cancelada no prazo legal'
		,'90'
		,'99'
		,'99'
		)
		,(
		371
		,'PES'
		,'6.949'
		,'Troca em Garantia SI'
		,'41'
		,'53'
		,'08'
		)
		,(
		372
		,'PES'
		,'5.949'
		,'Troca em Garantia SI'
		,'41'
		,'53'
		,'08'
		)
		,(
		373
		,'PSP'
		,'5.949'
		,'Troca em Garantia SI'
		,'41'
		,'53'
		,'08'
		)
		,(
		374
		,'PSP'
		,'6.949'
		,'Troca em Garantia SI'
		,'41'
		,'53'
		,'08'
		)
		,(
		375
		,'PES'
		,'1.949'
		,'Outras Entradas BN'
		,'00'
		,'00'
		,'74'
		)
		,(
		376
		,'PES'
		,'2.949'
		,'Outras Entradas BN'
		,'00'
		,'00'
		,'74'
		)
		,(
		377
		,'PSP'
		,'1.949'
		,'Outras Entradas (Bonifica��o)'
		,'00'
		,'00'
		,'74'
		)
		,(
		378
		,'PSP'
		,'2.949'
		,'Outras Entradas (Bonifica��o)'
		,'00'
		,'00'
		,'74'
		)
		,(
		381
		,'PSP'
		,'6.949'
		,'Simples Remessa - Sinistro'
		,'00'
		,'50'
		,'01'
		)
		,(
		382
		,'PES'
		,'5.949'
		,'Estorno de NF-e n�o cancelada no prazo legal'
		,'90'
		,'99'
		,'99'
		)
		,(
		383
		,'PSP'
		,'1.949'
		,'Retorno de Amostra p/ Teste'
		,'41'
		,'03'
		,'08'
		)
		,(
		385
		,'PSC'
		,'1.000'
		,'Entradas ou Aquisi��es de Servi�o'
		,'41'
		,'03'
		,'07'
		)
		,(
		386
		,'PSC'
		,'2.000'
		,'Entradas ou Aquisi��es de Servi�o'
		,'41'
		,'03'
		,'07'
		)
		,(
		387
		,'PSC'
		,'1.202'
		,'Devolu��o de Venda'
		,'00'
		,'00'
		,'99'
		)
		,(
		388
		,'PSC'
		,'2.202'
		,'Devolu��o de Venda'
		,'00'
		,'00'
		,'99'
		)
		,(
		389
		,'PSC'
		,'1.353'
		,'Transporte Rodovi�rio'
		,'00'
		,'03'
		,'07'
		)
		,(
		390
		,'PSC'
		,'2.353'
		,'Transporte Rodovi�rio'
		,'00'
		,'03'
		,'07'
		)
		,(
		393
		,'PSC'
		,'1.556'
		,'Compra p/ Uso ou Consumo'
		,'90'
		,'49'
		,'07'
		)
		,(
		394
		,'PSC'
		,'2.556'
		,'Compra p/ Uso ou Consumo'
		,'90'
		,'49'
		,'07'
		)
		,(
		395
		,'PSC'
		,'1.102'
		,'Compra p/ Revenda'
		,'00'
		,'00'
		,'08'
		)
		,(
		396
		,'PES'
		,'6.403'
		,'Complemento de ICMS ST'
		,'10'
		,'99'
		,'99'
		)
		,(
		397
		,'PES'
		,'6.108'
		,'Estorno de NF-e n�o cancelada no prazo legal'
		,'00'
		,'50'
		,'99'
		)
		,(
		398
		,'PSP'
		,'3.949'
		,'Entrada p/ Garantia'
		,'00'
		,'00'
		,'08'
		)
		,(
		399
		,'PSP'
		,'5.914'
		,'Remessa p/ Exposi��o ou Feira'
		,''
		,''
		,''
		)
		,(
		400
		,'PES'
		,'1.949'
		,'Estorno de NF-e n�o cancelada no prazo legal'
		,'90'
		,'49'
		,'99'
		)
		,(
		401
		,'PSP'
		,'1.949'
		,'Estorno de NF-e N�o Cancelada No Prazo Legal'
		,'90'
		,'49'
		,'99'
		)
		,(
		402
		,'PSP'
		,'5.914'
		,'Remessa p/ Exposi��o ou Feira'
		,'40'
		,'55'
		,'08'
		)
		,(
		403
		,'PES'
		,'2.932'
		,'Transporte Rodovi�rio'
		,'00'
		,'03'
		,'07'
		)
		,(
		404
		,'PES'
		,'1.932'
		,'Transporte Rodovi�rio'
		,'00'
		,'03'
		,'07'
		)
		,(
		405
		,'PES'
		,'5.927'
		,'Baixa de estoque - Uso Consumo'
		,'20'
		,'50'
		,'08'
		)
		,(
		406
		,'PES'
		,'5.927'
		,'Baixa de estoque - Perda / Contagem'
		,'20'
		,'50'
		,'08'
		)
		,(
		407
		,'PSP'
		,'5.927'
		,'Baixa de estoque - Perda/ Contagem'
		,'41'
		,'99'
		,'49'
		)
		,(
		408
		,'PES'
		,'5.927'
		,'Baixa de estoque - Garantia'
		,'20'
		,'50'
		,'08'
		)
		,(
		409
		,'PES'
		,'5.927'
		,'Baixa de estoque - Obsoletos'
		,'20'
		,'50'
		,'08'
		)
		,(
		410
		,'PES'
		,'5.927'
		,'Baixa de estoque - Desmembramento'
		,'20'
		,'50'
		,'08'
		)
		,(
		411
		,'PES'
		,'5.927'
		,'Baixa de estoque - Perda Importa��o'
		,'20'
		,'50'
		,'08'
		)
		,(
		413
		,'PSP'
		,'5.927'
		,'Baixa de estoque - Obsoletos'
		,'41'
		,'99'
		,'49'
		)
		,(
		414
		,'PSP'
		,'5.927'
		,'Baixa de estoque - Desmembramento'
		,'41'
		,'99'
		,'49'
		)
		,(
		415
		,'PSP'
		,'5.927'
		,'Baixa de estoque - Perda Importa��o'
		,'41'
		,'99'
		,'49'
		)
		,(
		416
		,'PES'
		,'5.102'
		,'Revenda (FS)'
		,'40'
		,'52'
		,'07'
		)
		,(
		417
		,'PES'
		,'6.102'
		,'Revenda (FS)'
		,'40'
		,'52'
		,'07'
		)
		,(
		418
		,'PSC'
		,'5.905'
		,'Rem. p/ Armazem'
		,'50'
		,'55'
		,'08'
		)
		,(
		419
		,'PSC'
		,'1.907'
		,'Ret Mercadoria Rem p/ Dep�sito Fech ou Armz Geral'
		,'50'
		,'05'
		,'74'
		)
		,(
		420
		,'PSC'
		,'5.106'
		,'Revenda'
		,'00'
		,'50'
		,'01'
		)
		,(
		421
		,'PSC'
		,'6.106'
		,'Revenda'
		,'00'
		,'50'
		,'01'
		)
		,(
		422
		,'PSC'
		,'6.403'
		,'Venda Adquirida de Terceiros c/ ST - (Substituto)'
		,'10'
		,'50'
		,'01'
		)
		,(
		423
		,'PSC'
		,'5.403'
		,'Venda Adquirida de Terceiros c/ ST - (Substituto)'
		,'10'
		,'50'
		,'01'
		)
		,(
		424
		,'PSC'
		,'6.108'
		,'Revenda p/ Consumo (PF)'
		,'00'
		,'50'
		,'01'
		)
		,(
		426
		,'PSC'
		,'5.106'
		,'Revenda p/ Consumo (PJ)'
		,'00'
		,'50'
		,'01'
		)
		,(
		427
		,'PSC'
		,'6.106'
		,'Revenda p/ Consumo (PJ)'
		,'00'
		,'50'
		,'01'
		)
		,(
		428
		,'PSC'
		,'6.110'
		,'Venda Adq. Terc. SUFRAMA (IPI)'
		,'00'
		,'55'
		,'01'
		)
		,(
		429
		,'PSC'
		,'6.110'
		,'Venda Adq. Terc. SUFRAMA Consumo'
		,'00'
		,'52'
		,'06'
		)
		,(
		430
		,'PSC'
		,'3.102'
		,'Importa��o p/ Revenda'
		,'00'
		,'00'
		,'01'
		)
		,(
		431
		,'PSP'
		,'2.949'
		,'Estorno de NF-e N�o Cancelada No Prazo Legal'
		,'90'
		,'49'
		,'98'
		)
		,(
		432
		,'PSC'
		,'1.411'
		,'Devolu��o de Venda c/ ST'
		,'10'
		,'00'
		,'50'
		)
		,(
		433
		,'PSC'
		,'2.411'
		,'Devolu��o de venda c/ ST'
		,'10'
		,'00'
		,'50'
		)
		,(
		434
		,'PSC'
		,'5.405'
		,'Venda Adquirida de Terceiros c/ ST - (Substitu�do)'
		,'60'
		,'50'
		,'01'
		)
		,(
		435
		,'PSC'
		,'1.949'
		,'Entrada de Importa��o por Conta e Ordem'
		,'00'
		,'00'
		,'08'
		)
		,(
		436
		,'PSC'
		,'2.152'
		,'Entrada Transf. p/ Comercializa��o'
		,'00'
		,'00'
		,'07'
		)
		,(
		437
		,'PSC'
		,'2.557'
		,'Entrada Transf. p/ Uso e Consumo'
		,'00'
		,'00'
		,'08'
		)
		,(
		438
		,'PSC'
		,'6.152'
		,'Transf. de Mercadorias'
		,'00'
		,'50'
		,'08'
		)
		,(
		439
		,'PSC'
		,'6.557'
		,'Transf. p/ Uso e Consumo'
		,'00'
		,'50'
		,'08'
		)
		,(
		440
		,'PSC'
		,'5.106'
		,'Revenda (Aduirente Industrial)'
		,'00'
		,'50'
		,'01'
		)
		,(
		441
		,'PSC'
		,'6.106'
		,'Revenda (Adquirente Industrial)'
		,'00'
		,'50'
		,'01'
		)
		,(
		442
		,'PSP'
		,'5.102'
		,'Revenda p/ Consumo (PF)'
		,'00'
		,'50'
		,'01'
		)
		,(
		443
		,'PES'
		,'5.102'
		,'REVENDA p/ Consumo (PF)'
		,'00'
		,'50'
		,'01'
		)
		,(
		444
		,'PSC'
		,'5.106'
		,'Revenda p/ Consumo (PF)'
		,'00'
		,'50'
		,'01'
		)
		,(
		445
		,'PSC'
		,'6.949'
		,'Emprestimo de Mercadoria'
		,'00'
		,'53'
		,'08'
		)
		,(
		446
		,'PSC'
		,'5.949'
		,'Emprestimo de Mercadoria'
		,'00'
		,'53'
		,'08'
		)
		,(
		447
		,'PES'
		,'6.110'
		,'Venda Adq. Terc. SUFRAMA (IPI e P/C)'
		,'00'
		,'55'
		,'06'
		)
		,(
		448
		,'PSP'
		,'6.110'
		,'Venda Adq. Terc. SUFRAMA (IPI e P/C)'
		,'00'
		,'55'
		,'06'
		)
		,(
		449
		,'PES'
		,'6.102'
		,'Venda Adq. Terc. SUFRAMA (SD)'
		,'00'
		,'50'
		,'01'
		)
		,(
		450
		,'PSP'
		,'6.102'
		,'Venda Adq. Terc. SUFRAMA (SD)'
		,'00'
		,'50'
		,'01'
		)
		,(
		451
		,'PSP'
		,'2.949'
		,'Entrada p/ Brindes'
		,'00'
		,'00'
		,'08'
		)
		,(
		452
		,'PSP'
		,'1.949'
		,'Entrada p/ Brindes'
		,'00'
		,'00'
		,'08'
		)
		,(
		453
		,'PSC'
		,'5.910'
		,'Remessa em Bonifica��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		454
		,'PSC'
		,'6.910'
		,'Remessa em Bonifica��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		455
		,'PSC'
		,''
		,'sufra'
		,''
		,''
		,''
		)
		,(
		456
		,'PSC'
		,'6.110'
		,'Venda Adq. Terc. SUFRAMA (IPI e P/C)'
		,'00'
		,'52'
		,'06'
		)
		,(
		459
		,'PSC'
		,'6.102'
		,'Venda Adq. Terc. SUFRAMA (SD)'
		,'00'
		,'50'
		,'01'
		)
		,(
		460
		,'PSP'
		,'3.949'
		,'Importa��o p/ Produtos de Homologa��o'
		,'41'
		,'03'
		,'70'
		)
		,(
		463
		,'PSC'
		,'5.927'
		,'Baixa de estoque - Perda Importa��o'
		,'00'
		,'53'
		,'08'
		)
		,(
		464
		,'PSC'
		,'5.927'
		,'Baixa de estoque - Desmembramento'
		,'00'
		,'53'
		,'08'
		)
		,(
		465
		,'PSC'
		,'5.927'
		,'Baixa de estoque - Garantia'
		,'00'
		,'53'
		,'08'
		)
		,(
		466
		,'PSC'
		,'5.927'
		,'Baixa de estoque - Roubo/ Furto'
		,'00'
		,'53'
		,'08'
		)
		,(
		467
		,'PSC'
		,'5.927'
		,'Baixa de estoque - Uso Consumo'
		,'00'
		,'53'
		,'08'
		)
		,(
		468
		,'PSC'
		,'5.927'
		,'Baixa de estoque - Perda / Contagem'
		,'00'
		,'53'
		,'08'
		)
		,(
		469
		,'PSC'
		,'5.927'
		,'Baixa de estoque - Obsoletos'
		,'00'
		,'53'
		,'08'
		)
		,(
		470
		,'PES'
		,'6.910'
		,'Remessa em Bonifica��o (AJ)'
		,'00'
		,'50'
		,'08'
		)
		,(
		471
		,'PSC'
		,'3.949'
		,'Importa��o p/ Garantia - Pagamento Invoice'
		,'90'
		,'49'
		,'99'
		)
		,(
		472
		,'PSC'
		,''
		,'Compra para Comercializa��o'
		,'00'
		,'00'
		,'70'
		)
		,(
		473
		,'PSC'
		,'3.102'
		,'Compra para Comercializa��o'
		,'00'
		,'00'
		,'70'
		)
		,(
		474
		,'PSC'
		,'3.102'
		,'Complemento de Importa��o'
		,'00'
		,'00'
		,'01'
		)
		,(
		475
		,'PSC'
		,'5.949'
		,'Troca em Garantia SI'
		,'41'
		,'53'
		,'08'
		)
		,(
		476
		,'PSC'
		,'6.949'
		,'Troca em Garantia SI'
		,'41'
		,'53'
		,'08'
		)
		,(
		477
		,'PSP'
		,'7.949'
		,'Outras Sa�das de Mercadoria/Presta��o de Servi�o'
		,'90'
		,'53'
		,'49'
		)
		,(
		478
		,'PSC'
		,'1.949'
		,'Outras Entradas'
		,'90'
		,'49'
		,'98'
		)
		,(
		479
		,'PSC'
		,'5.916'
		,'Retorno Mercadoria Recebida Conserto/Reparo'
		,'40'
		,'53'
		,'08'
		)
		,(
		480
		,'PSC'
		,'6.916'
		,'Retorno Mercadoria Recebida Conserto/Reparo'
		,'40'
		,'53'
		,'08'
		)
		,(
		481
		,'PSC'
		,'2.949'
		,'Outras Entradas (Bonifica��o)'
		,'00'
		,'00'
		,'74'
		)
		,(
		482
		,'PSC'
		,'1.949'
		,'Outras Entradas BN'
		,'00'
		,'00'
		,'74'
		)
		,(
		483
		,'PSP'
		,'1.949'
		,'Retorno de Emprestimo de Mercadoria'
		,'00'
		,'03'
		,'08'
		)
		,(
		484
		,'PSP'
		,'2.949'
		,'Retorno de Emprestimo de Mercadoria'
		,'00'
		,'03'
		,'08'
		)
		,(
		485
		,'PES'
		,'1.949'
		,'Retorno de Emprestimo de Mercadoria'
		,'00'
		,'03'
		,'08'
		)
		,(
		486
		,'PES'
		,'2.949'
		,'Retorno de Emprestimo de Mercadoria'
		,'00'
		,'03'
		,'08'
		)
		,(
		487
		,'PSC'
		,'1.949'
		,'Retorno de Emprestimo de Mercadoria'
		,'00'
		,'03'
		,'08'
		)
		,(
		488
		,'PSC'
		,'2.949'
		,'Retorno de Emprestimo de Mercadoria'
		,'00'
		,'03'
		,'08'
		)
		,(
		489
		,'PSC'
		,'5.910'
		,'Remessa de Doa��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		490
		,'PSC'
		,'6.910'
		,'Remessa de Doa��o'
		,'00'
		,'50'
		,'08'
		)
		,(
		491
		,'PSC'
		,'5.905'
		,'Complemento de Rem. p/ Armazem'
		,'50'
		,'55'
		,'08'
		)
		,(
		492
		,'PSC'
		,'5.949'
		,'Simples Remessa - Sinistro'
		,'00'
		,'50'
		,'01'
		)
		,(
		493
		,'PSC'
		,'6.949'
		,'Simples Remessa - Sinistro'
		,'00'
		,'50'
		,'01'
		)
		,(
		494
		,'PSC'
		,'6.949'
		,'Estorno de NF-e n�o cancelada no prazo legal'
		,'41'
		,'00'
		,'01'
		)
		,(
		495
		,'PSP'
		,'6.949'
		,'Estorno de NF-e n�o cancelada no prazo legal'
		,'00'
		,'00'
		,'01'
		)
		,(
		496
		,'PSP'
		,'7.949'
		,'Estorno de NF-e n�o cancelada no prazo legal'
		,'00'
		,'00'
		,'01'
		)
		,(
		497
		,'PSC'
		,'6.110'
		,'Complemento do ICMS'
		,''
		,''
		,''
		)
		,(
		498
		,'PSC'
		,'1.907'
		,'Estorno de Rem. p/ Armaz.'
		,'50'
		,'05'
		,'74'
		)
		,(
		499
		,'PSP'
		,'6.949'
		,'Garantia'
		,'41'
		,'53'
		,'08'
		)
		,(
		500
		,'PSP'
		,'3.949'
		,'Importa��o p/ Amostras'
		,'41'
		,'03'
		,'70'
		)
		,(
		502
		,'PSP'
		,'5.949'
		,'Remessa Uso/Consumo p/ Uso Fora do Estabelecimento'
		,'50'
		,'53'
		,'08'
		)
		,(
		503
		,'PSP'
		,'6.554'
		,'Remessa de Ativo p/ Uso Fora do Estabelecimento'
		,'50'
		,'53'
		,'08'
		)
		,(
		504
		,'PSP'
		,'6.949'
		,'Remessa Uso/Consumo p/ Uso Fora do Estabelecimento'
		,'50'
		,'53'
		,'08'
		)
		,(
		505
		,'PSC'
		,'1.926'
		,'Entrada decorrente de desagrega��o'
		,'40'
		,'49'
		,'99'
		)
		,(
		506
		,'PSC'
		,'5.926'
		,'Remessa p/ Desagrega��o'
		,'40'
		,'49'
		,'99'
		)
		,(
		507
		,'PSC'
		,'5.206'
		,'Anula��o de Valor Relativo � Servi�o de Transporte'
		,'90'
		,'53'
		,'08'
		)
		,(
		508
		,'PSP'
		,'3.949'
		,'Importa��o p/ Garantia - Sem Pagamento Invoice'
		,'41'
		,'03'
		,'70'
		)
		,(
		509
		,'PSC'
		,'1.949'
		,'Estorno de NF-e N�o Cancelada No Prazo Legal'
		,'90'
		,'49'
		,'98'
		)
		,(
		510
		,'PSP'
		,'5.949'
		,'Retorno em Garantia'
		,'90'
		,'99'
		,'99'
		)
		,(
		511
		,'PSP'
		,'2.933'
		,'Aquisi��o de Servi�o Tributado pelo ISSQN'
		,'90'
		,'49'
		,'70'
		)
		,(
		512
		,'PSP'
		,'7.949'
		,'Amostra p/ Teste (Exterior)'
		,'41'
		,'53'
		,'49'
		)
		,(
		514
		,'PSC'
		,''
		,'Complemento Retorno de Emprestimo de Mercadoria'
		,'90'
		,'49'
		,'99'
		)
		,(
		515
		,'PSC'
		,'1.949'
		,'Complemento - Retorno de Emprestimo de Mercadoria'
		,'90'
		,'03'
		,'98'
		)
		,(
		516
		,'PSC'
		,'5.914'
		,'Remessa p/ Exposi��o ou Feira'
		,'40'
		,'55'
		,'08'
		)
		,(
		517
		,'PSC'
		,'1.914'
		,'Retorno de Remessa p/ Exposi��o/Feira'
		,'40'
		,'05'
		,'08'
		)
		,(
		518
		,'PSC'
		,'2.914'
		,'Retorno de Remessa p/ Exposi��o/Feira'
		,'00'
		,'03'
		,'08'
		)
		,(
		519
		,'PSC'
		,'6.914'
		,'Remessa p/ Exposi��o ou Feira'
		,'00'
		,'53'
		,'08'
		)
		,(
		520
		,'PSP'
		,'3.353'
		,'Transporte Rodovi�rio'
		,'90'
		,'03'
		,'70'
		)
		,(
		521
		,'PSC'
		,'3.353'
		,'Transporte Rodovi�rio'
		,'90'
		,'03'
		,'70'
		)
		,(
		522
		,'PES'
		,'3.353'
		,'Transporte Rodovi�rio'
		,'90'
		,'03'
		,'70'
		)
		,(
		523
		,'PSC'
		,'1.949'
		,'Outras Entradas - Importa��o'
		,'90'
		,'49'
		,'99'
		)
		,(
		524
		,'PSP'
		,'1.949'
		,'Outras Entradas - Importa��o'
		,'90'
		,'49'
		,'99'
		)
		,(
		526
		,'PSP'
		,'6.556'
		,'Devolu��o de Compra de Material de Uso ou Consumo'
		,'00'
		,'53'
		,'08'
		)
		,(
		527
		,'PSC'
		,'5.905'
		,'Estorno de Rem. p/ Armazem'
		,'41'
		,'53'
		,'08'
		)
		,(
		528
		,'PSC'
		,'1.551'
		,'Compra de Bem p/ Ativo Imobilizado'
		,'90'
		,'49'
		,'07'
		)
		,(
		529
		,'PSC'
		,'2.551'
		,'Compra de Bem p/ Ativo Imobilizado'
		,'90'
		,'49'
		,'07'
		)
		,(
		530
		,'PSP'
		,'2.949'
		,'Retorno de Amostra p/ Teste'
		,'41'
		,'03'
		,'08'
		)
		,(
		531
		,'PES'
		,'2.949'
		,'Retorno de Amostra p/ Teste'
		,'41'
		,'03'
		,'08'
		)
		,(
		532
		,'PES'
		,'1.949'
		,'Retorno de Amostra p/ Teste'
		,'41'
		,'03'
		,'08'
		)
		,(
		533
		,'PSC'
		,'1.949'
		,'Retorno de Amostra p/ Teste'
		,'41'
		,'03'
		,'08'
		)
		,(
		534
		,'PSC'
		,'2.949'
		,'Retorno de Amostra p/ Teste'
		,'41'
		,'03'
		,'08'
		)
		,(
		536
		,'PSC'
		,'1.407'
		,'Compra de Bem p/ Uso ou Consumo c/ ST'
		,'60'
		,'49'
		,'07'
		)
		,(
		537
		,'PSC'
		,'5.114'
		,'Venda de Consigna��o'
		,'41'
		,'53'
		,'01'
		)
		,(
		538
		,'PSC'
		,'1.949'
		,'Outras entradas'
		,'90'
		,'49'
		,'98'
		)
		,(
		539
		,'PSC'
		,'2.949'
		,'Outras entradas'
		,'90'
		,'49'
		,'98'
		)
		,(
		540
		,'PSP'
		,'6.553'
		,'Devolu��o de compra de ativo imobilizado'
		,'41'
		,'53'
		,'08'
		)
		,(
		541
		,'PSP'
		,'5.914'
		,'Remessa p/ Exposi��o ou Feira - Estrutura'
		,'40'
		,'55'
		,'08'
		)
		,(
		542
		,'PSP'
		,'6.914'
		,'Remessa p/ Exposi��o ou Feira - Estrutura'
		,'00'
		,'53'
		,'08'
		)
		,(
		543
		,'PSP'
		,'1.914'
		,'Retorno de Remessa p/ Exposi��o/Feira - Estrutura'
		,'40'
		,'05'
		,'08'
		)
		,(
		544
		,'PSP'
		,'2.914'
		,'Retorno de Remessa p/ Exposi��o/Feira - Estrutura'
		,'00'
		,'03'
		,'08'
		)
		,(
		545
		,'PSC'
		,'2.949'
		,'Outras Entradas (Material Promocional)'
		,'90'
		,'49'
		,'98'
		)
		,(
		546
		,'PSC'
		,'2.949'
		,'Estorno Troca em Garantia SI'
		,'41'
		,'03'
		,'74'
		)
		,(
		547
		,'PSC'
		,'1.949'
		,'Estorno Troca em Garantia SI'
		,'41'
		,'03'
		,'08'
		)
		,(
		548
		,'PSC'
		,'1.949'
		,'Entrada p/ Garantia'
		,'90'
		,'03'
		,'08'
		)
		,(
		549
		,'PSC'
		,'2.949'
		,'Entrada p/ Garantia'
		,'90'
		,'03'
		,'08'
		)
		,(
		550
		,'PSC'
		,'5.949'
		,'Troca em Garantia'
		,'00'
		,'53'
		,'08'
		)
		,(
		551
		,'PSC'
		,'6.949'
		,'Troca em Garantia'
		,'00'
		,'53'
		,'08'
		)
		,(
		552
		,'PSC'
		,'1.906'
		,'Ret Mercadoria Rem p/ Dep�sito Fech ou Armz Geral'
		,'50'
		,'05'
		,'74'
		)
		,(
		553
		,'PSC'
		,'6.152'
		,'Transf. de Mercadorias (P/ An�lise)'
		,'00'
		,'50'
		,'08'
		)
		,(
		554
		,'PSC'
		,'2.102'
		,'Compra p/ Revenda'
		,'00'
		,'03'
		,'98'
		)
		,(
		555
		,'PSP'
		,'3.102'
		,'Importa��o p/ Garantia - Pagamento Invoice'
		,'00'
		,'00'
		,'01'
		)
		,(
		556
		,'PSC'
		,'6.106'
		,'Venda Adq. Terc. SUFRAMA (IPI) Fora da ZF'
		,'00'
		,'55'
		,'01'
		)
		,(
		557
		,'PSC'
		,'6.106'
		,'Venda Adq. Terc. SUFRAMA (IPI e P/C) Fora da ZF'
		,'00'
		,'52'
		,'06'
		)
		,(
		558
		,'PSP'
		,'6.102'
		,'Venda Adq. Terc. SUFRAMA (IPI) Fora da ZF'
		,'00'
		,'55'
		,'01'
		)
		,(
		559
		,'PSP'
		,'6.102'
		,'Venda Adq. Terc. SUFRAMA (IPI e P/C) Fora da ZF'
		,'00'
		,'55'
		,'06'
		)
		,(
		560
		,'PSC'
		,'2.932'
		,'Transporte Rodovi�rio'
		,'00'
		,'03'
		,'07'
		)
		,(
		561
		,'PSC'
		,'1.932'
		,'Transporte Rodovi�rio'
		,'00'
		,'03'
		,'07'
		)
		,(
		562
		,'PSP'
		,'1.932'
		,'Transporte Rodovi�rio'
		,'00'
		,'03'
		,'07'
		)
		,(
		563
		,'PSP'
		,'2.932'
		,'Transporte Rodovi�rio'
		,'00'
		,'03'
		,'07'
		)
	) AS tmp(ID_NatOper, Fil_NatOper, CFOP_NatOper, Descr_NatOper, STICMS_NatOper, STIPI_NatOper, STPC_NatOper)

SET IDENTITY_INSERT tblNatOp OFF
