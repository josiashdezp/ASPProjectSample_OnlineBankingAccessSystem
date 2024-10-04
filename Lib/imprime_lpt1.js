///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////                                    Criado por : Flávio Theruo Kaminisse                                   ////////
////////                                         email: flavio@japs.etc.br                                         ////////
////////                                         url: http://www.japs.etc.br                                       ////////
////////                                          Data Criação : 12/11/2005                                        ////////
////////                                                                                                           ////////
////////                                       - Compatível Internet Explorer.                                     ////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/**************************************************************************************************************************

function imprime_lpt1(arg1)
Descricao:  Funcao responsavel pela impressao diretamente na porta lpt1.

Requisitos: Internet Explorer superior que 5.0, mas nao testado no 5.5, testado e funcionando no 6.0.
			Nao testado em outros browsers

			****** IMPORTANTISSIMO ******
			Se estas alteracoes nao forem feitas, a impressao nao ocorrera corretamente.
			Alteracao das configuracoes de seguranca do navegador:
			Ferramentas -> Opcoes de Internet -> Seguranca ->
			Intranet local -> Nivel Personalizado ->
			Iniciar e executar scripts de controle ActiveX nao marcados como seguro -> Ativar
			Caso nao funcione:
			Internet -> Nivel Personalizado ->
			Iniciar e executar scripts de controle ActiveX nao marcados como seguro -> Ativar

Entrada:    Linha que sera impressa na autenticacao ou o comprovante de pagamento.

Saida:      Impressao em impressora fiscal do argumento de entrada.

**************************************************************************************************************************/

function Print_To_Port(arg1) {
	
	alert(arg1);
	//Cria objeto para manipulacao de arquivos no cliente.
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	//Verifica a existencia do arquivo de impressao antigo e o deleta;
	if ( fso.FileExists("c:\\temp\\imprime.prn") ) {
		fso.DeleteFile("c:\\temp\\imprime.prn");
	} //if
	//Verifica a nao existencia do arquivo responsavel pela impressao na impressora bematech.
	if ( !(fso.FileExists("c:\\temp\\imprime.bat")) ) {
		//Cria o arquivo imprime.bat, escreve o comando responsavel pela impressao e fecha o arquivo.
		var b = fso.CreateTextFile("c:\\temp\\imprime.bat", true);
		b.WriteLine("type c:\\temp\\imprime.prn > prn");
		b.Close();
	} //if
	//Cria o arquivo imprime.prn, escreve todo o texto para ser impresso e fecha o arquivo.
	var a = fso.CreateTextFile("c:\\temp\\imprime.prn", true);
	a.WriteLine(arg1);
	a.Close();
	//Cria um objeto para execucao de um programa no computador do cliente.
	//var WshShell = new ActiveXObject("WScript.Shell");
	//Executa o arquivo responsavel pela impressao do arquivo imprime.prn.
	//var oExec = WshShell.Exec("c:\\temp\\imprime.bat");
} //Fim do imprime_lpt1