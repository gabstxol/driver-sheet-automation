//--Automação de Planilha de Motoristas --//

//PROBLEMA DA ÁREA//

/*Na área determinada da empresa havia uma dificuldade enorme, pois os motoristas demoravam muito pra informar seus dados a portaria da entrada;
e depois passar os numeros das notas fiscais(eram várias notas fiscais por se tratar de muito material entregue.

//SOLUÇÃO//
então pensei em implementar formularios distintos:
1-Registro de todos os dados do motorista e a primeira nota fiscal
2-RG e Nota Fiscal
Depois disso adicionei um fluxo do Power Automate no qual separava os dados, enviando para a planilha os dois formularios.
//CÓDIGO//
 O código consistem em, pegar os dados da planilha, e depois coloca-los em uma lista, e verificar se tem algum motorista que tenha o rg igual,
 então ele verifica desses dois qual tem o Nome nulo ou vazio, e depois preenche os campos faltosos na linha que somente contém o RG e a Nota Fiscal, e logo após os dados
 são enviados para a planilha.
*/
