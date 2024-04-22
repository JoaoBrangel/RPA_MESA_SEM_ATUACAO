/*
DATA: 17/04/2024
NOME DO JOB: [ITAU PJ] - MESA SEM ATUACAO INSERIR
AUTOR: JOÃO LUIZ - TI
SOLICITANTE: MICHELI ENKIN
MOTIVO: AUTOMATIZAR O PROCESSO, ESTAVAMOS FAZENDO MANUALMENTE.

TUTORIAL: extraia o arquivo do .zip e deixe ele com a nomeclatura ( Mesa_sem_atuacao_data.xlsx ) e execute o job.


OBS: O SCRIPT ELE BUSCA A PASTA DO EDI ANTES DAS 14:00H NO MODELO (DIA_MES) E DEPOIS DAS 14:00H NO MODELO (DIA_MES_1)
*/


use safra;

DECLARE @DIRETORIO				VARCHAR(200)
DECLARE @COMANDO				VARCHAR(200)
DECLARE @ARQUIVOESCOLHIDO		VARCHAR(200)
DECLARE	@DIA					VARCHAR(20)		=	DAY(GETDATE())
DECLARE	@ANO					VARCHAR(20)		=	YEAR(GETDATE())
DECLARE	@MES					VARCHAR(20)		
DECLARE @PASTA_DIA_MES			VARCHAR(200)	=	(select replace(convert(varchar(5), getdate(), 103), '/', '_'))
DECLARE @COPY					VARCHAR(200)
DECLARE @MOVE					VARCHAR(255)
DECLARE @CAMINHO_PROCESSADOS	VARCHAR(255)	=   '\\192.168.0.46\import_cobweb\ITAU_MESA_SEM_ATUACAO\PROCESSADOS'
DECLARE @DESTINO				VARCHAR(200)	=	'\\192.168.0.46\import_cobweb\ITAU_MESA_SEM_ATUACAO' 
DECLARE @RENAME					VARCHAR(255);

SET		@MES	=	(
						CASE 
							WHEN	MONTH(GETDATE())	=	1	THEN	'01 - JANEIRO'
							WHEN	MONTH(GETDATE())	=	2	THEN	'02 - FEVEREIRO'
							WHEN	MONTH(GETDATE())	=	3	THEN	'03 - MARÇO'
							WHEN	MONTH(GETDATE())	=	4	THEN	'04 - ABRIL'
							WHEN	MONTH(GETDATE())	=	5	THEN	'05 - MAIO'
							WHEN	MONTH(GETDATE())	=	6	THEN	'06 - JUNHO'
							WHEN	MONTH(GETDATE())	=	7	THEN	'07 - JULHO'
							WHEN	MONTH(GETDATE())	=	8	THEN	'08 - AGOSTO'
							WHEN	MONTH(GETDATE())	=	9	THEN	'09 - SETEMBRO'
							WHEN	MONTH(GETDATE())	=	10	THEN	'10 - OUTUBRO'
							WHEN	MONTH(GETDATE())	=	11	THEN	'11 - NOVEMBRO'
							WHEN	MONTH(GETDATE())	=	12	THEN	'12 - DEZEMBRO'
						END
					)

IF(convert(varchar(5),GETDATE(),108)) < '14:00' 
begin
		SET	@DIRETORIO	=	'"\\192.168.0.48\EDI\' + @ano + '\' + @MES + '\' + @PASTA_DIA_MES
		SET	@COMANDO	=	'dir ' + @DIRETORIO + '" /b'
end
else
begin
		SET	@DIRETORIO	=	'"\\192.168.0.48\EDI\' + @ano + '\' + @MES + '\' + @PASTA_DIA_MES + '_' + '1'
		SET	@COMANDO	=	'dir ' + @DIRETORIO + '" /b'

end

-- ============================================================================================================
		
DROP TABLE IF EXISTS #ARQUIVOS

CREATE TABLE #ARQUIVOS
(	NOMEARQUIVO VARCHAR (200)
)

INSERT INTO #ARQUIVOS
EXEC xp_cmdshell	@comando 


SET @ARQUIVOESCOLHIDO = (
							SELECT 
									NOMEARQUIVO 
							FROM	#ARQUIVOS	
							WHERE	NOMEARQUIVO LIKE	'%Mesa_sem_atuacao%'
							and		NOMEARQUIVO LIKE	'%' + @DIA + '%' + SUBSTRING(@MES, 1, 2) + '%' + @ANO + '%'
						)
print @ARQUIVOESCOLHIDO

-- ============================================================================================================

begin try
		-- Comando para copiar o arquivo usando o comando COPY do Windows
		SET @COPY = 'COPY ' + @DIRETORIO + '\'+ @ARQUIVOESCOLHIDO + '"  "' + @DESTINO +'"';

		--PRINT @COPY 
		-- Executar o comando usando xp_cmdshell
		EXEC xp_cmdshell @COPY;

-- ============================================================================================================ TABELA PARA PEGAR O ARQUIVO
DECLARE @nome_arquivo		varchar(250)	= (@DESTINO + '\' + @ARQUIVOESCOLHIDO)
DECLARE @OpenRowSet			varchar(max);

print	@nome_arquivo
		


					Set		@OpenRowSet = 
					'DROP TABLE if exists MESA_SEM_ATAUCAO_IMPORT
					SELECT 
					*
					into			MESA_SEM_ATAUCAO_IMPORT
					FROM OPENROWSET	(''Microsoft.ACE.OLEDB.12.0'', 
									''Excel 12.0;Database='+ @nome_arquivo +';'',
									''SELECT * FROM [Planilha1$]'')'

Exec (@OpenRowSet)

-- ============================================================================================================ TABELA TRATATIVA


	insert INTO TBL_MESA_SEM_ATUACAO_INSERIR 
				SELECT 
				RIGHT(CNPJ_COMPLETO, 15)		AS	CNPJ_COMPLETO,
				''								AS	NOME_CLI,
				''								AS	ATR_CLI,
				''								AS	VLR_CA6_CLI,
				''								AS	FILA,
				Mesa_sem_atuacao				AS	Alcada_Mesa
				FROM	MESA_SEM_ATAUCAO_IMPORT
				where	CNPJ_COMPLETO			IS NOT NULL
				AND		Mesa_sem_atuacao				IS NOT NULL




SELECT * FROM TBL_MESA_SEM_ATUACAO_INSERIR
--SELECT * FROM TBL_MESA_SEM_ATUACAO_INSERIR_TEMP

-- ============================================================================================================

	DROP TABLE if exists #temp_email_qtde_entradas_gatilhos
	create table #temp_email_qtde_entradas_gatilhos
		(
			QTDADE_MESA_SEM_ATUACAO		varchar(100)not null,
		)

	INSERT INTO #temp_email_qtde_entradas_gatilhos (QTDADE_MESA_SEM_ATUACAO)
		SELECT 
				COUNT		(ai.CNPJ_COMPLETO) as [QTDE DE CASOS]
				FROM		TBL_MESA_SEM_ATUACAO_INSERIR			ai
				LEFT JOIN	TBL_MESA_SEM_ATUACAO					ma 
				ON			ai.CNPJ_COMPLETO						= ma.CNPJ_COMPLETO






-- ============================================================================================================


		
MERGE 
			TBL_MESA_SEM_ATUACAO AS Destino
		USING 
			TBL_MESA_SEM_ATUACAO_INSERIR AS ORIGEM ON (ORIGEM.CNPJ_COMPLETO = DESTINO.CNPJ_COMPLETO)
 
		-- Registro não existe no destino. Vamos inserir.
		WHEN NOT MATCHED THEN
			INSERT
			VALUES(ORIGEM.CNPJ_COMPLETO, ORIGEM.NOME_CLI, ATR_CLI, VLR_CA6_CLI, FILA, ALCADA_MESA)

		-- Registro existe no destino, mas, não existe na origem

		--WHEN NOT MATCHED BY SOURCE THEN
		--    DELETE
		;


-- ============================================================================================================


SET @MOVE = 'MOVE "' + @DESTINO + '\' + @ARQUIVOESCOLHIDO + '" "' + @CAMINHO_PROCESSADOS + '\' + @ARQUIVOESCOLHIDO + '"';
		--PRINT @COPY 
		EXEC xp_cmdshell @MOVE;	-- Executar o comando usando xp_cmdshell



DECLARE @COMPLEMENTO_NOME		VARCHAR(255);
		DECLARE @NOVO_NOME_ARQUIVO		VARCHAR(255);
		DECLARE @Renomear				VARCHAR(500);

		-- Remove os primeiros três caracteres
		SET @ARQUIVOESCOLHIDO = LEFT(@ARQUIVOESCOLHIDO, LEN(@ARQUIVOESCOLHIDO) - 5);

		-- Gera o complemento do nome com segundos e milissegundos
		SET @COMPLEMENTO_NOME = RIGHT('0'	+ CONVERT(VARCHAR(2), DATEPART(SECOND, GETDATE())), 2) + 
								RIGHT('000' + CONVERT(VARCHAR(4), DATEPART(MILLISECOND, GETDATE())), 3);

		-- Monta o novo nome do arquivo com o complemento de nome e a nova extensão
		SET @NOVO_NOME_ARQUIVO = @ARQUIVOESCOLHIDO + '_' + @COMPLEMENTO_NOME + '.xlsx';

		-- Monta o comando de renomear o arquivo
		SET @Renomear = 'REN "' + @CAMINHO_PROCESSADOS + '\' + @ARQUIVOESCOLHIDO + '.xlsx" "' + @NOVO_NOME_ARQUIVO + '"';

		-- Imprime o comando de renomear para verificação
		PRINT @Renomear;

		-- Executar o comando usando xp_cmdshell
		EXEC xp_cmdshell @Renomear;


-- ============================================================================================================
		
			DECLARE @para		VARCHAR(1000)	= '';
			DECLARE @assunto	VARCHAR(1000)	= 'Mesa sem atuação  - ' + FORMAT(GETDATE(), 'dd/MM/yyyy');
			DECLARE @mensagem	VARCHAR(MAX)	= '';
	
			SET @para += 'joao.reis@novaquest.com.br;';
			SET @para += 'vinicius@novaquest.com.br;';
			SET @para += 'micheli@novaquest.com.br';
			SET @para += 'sistemas@novaquest.com.br';
			SET @para += 'marcos.damasceno@novaquest.com.br';
			SET @para += 'victor.luis@novaquest.com.br';
			SET @para += 'mariuxa.tiburcio@novaquest.com.br';

			SET @mensagem += '<style type="text/css">';
			SET @mensagem += 'table, th, td {border: 1px solid black; border-collapse: collapse; padding: 0 5px 0 5px;}';
			SET @mensagem += 'p {font-size: 12pt;}';
			SET @mensagem += '</style>';
			SET @mensagem += '<style type="text/css">';
			SET @mensagem += 'table, th, td {border: 1px solid black; border-collapse: collapse; padding: 0 5px 0 5px;}';
			SET @mensagem += 'th {background-color: #FF7200;color:white}'; -- Estilo para o cabeçalho
			SET @mensagem += 'p {font-size: 12pt;}';
			SET @mensagem += '</style>';

			SET @mensagem += '<h3 style="text-align: center;">Mesa sem atuação importada em sistemas</h3>';
			SET @mensagem += '<h4 style="text-align: center;">Quantidades de casos importados</h4>';
			SET @mensagem += '</br>';
			SET @mensagem += '<table align="center" style="text-align: center;" >';
			SET @mensagem += '<tr>';
			SET @mensagem += '<th>MESA_SEM_ATUACAO</th>';
			SET @mensagem += '</tr>';
			SET @mensagem += (SELECT

								'<td>' +	QTDADE_MESA_SEM_ATUACAO + '</td>' 
							 FROM #temp_email_qtde_entradas_gatilhos
							 FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)');

			SET @mensagem += '</table>';
			SET @mensagem += '</br>';
			SET @mensagem += '<h5 style="text-align: center;">Nome do job: [ITAU PJ] - MESA SEM ATUACAO INSERIR </h5>';

			EXEC MSDB.DBO.SP_SEND_DBMAIL
				@recipients = @para,
				@subject = @assunto,
				@body = @mensagem,
				@body_format = 'HTML';

end try
begin CATCH
		
			DECLARE @para1		VARCHAR(1000)	= '';
			DECLARE @assunto1	VARCHAR(1000)	= 'Mesa sem atuação  - ERRO - TESTE ' + FORMAT(GETDATE(), 'dd/MM/yyyy');
			DECLARE @mensagem1	VARCHAR(MAX)	= '';

			SET @para1 += 'joao.reis@novaquest.com.br;';
			SET @para1 += 'vinicius@novaquest.com.br;';
			SET @para += 'micheli@novaquest.com.br';
			SET @para1 += 'sistemas@novaquest.com.br';

			SET @mensagem1 += '<style type="text/css">';
			SET @mensagem1 += 'table, th, td {border: 1px solid black; border-collapse: collapse; padding: 0 5px 0 5px;}'
			SET @mensagem1 += '</style>'
			SET @mensagem1 += '<h3 style="text-align: center;">Não foi possível importar o arquivo de mesa sem atuação.</h3>';
			SET @mensagem1 += '</br>';
			SET @mensagem1 += '<table align="center" style="text-align: center;" >';
			SET @mensagem1 += '<tr>';
			SET @mensagem1 += '<th>MESA_SEM_ATUACAO</th>';
			SET @mensagem1 += '</tr>';
			SET @mensagem1 += '</table>';
			SET @mensagem1 += '</br>';
			SET @mensagem1 += '<h5 style="text-align: center;">Nome do job: [ITAU PJ] - MESA SEM ATUACAO INSERIR </h5>';

			EXEC MSDB.DBO.SP_SEND_DBMAIL
				@recipients = @para1,
				@subject = @assunto1,
				@body = @mensagem1,
				@body_format = 'HTML';

END CATCH




