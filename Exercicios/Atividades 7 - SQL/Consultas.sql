CREATE DATABASE Lab;

----------------------------------------

DROP TABLE IF EXISTS livros;
CREATE TABLE livros (
  id_livro BIGINT NOT NULL ,
  nome VARCHAR(45) DEFAULT NULL,
  edicao NUMERIC(11) DEFAULT NULL,
  dt_expedicao DATE DEFAULT NULL,
  qtd_estoque NUMERIC(11) DEFAULT NULL,
  preco FLOAT DEFAULT NULL,
  PRIMARY KEY (id_livro),
);

INSERT INTO livros (id_livro, nome, edicao, dt_expedicao, qtd_estoque, preco) VALUES (3,'O nome do vento',2,'2009-06-23',100,41.27),(4,'O temor do sábio',1,'2011-11-20',77,45.63),(5,'Armas de Tortuga',1,'1999-04-28',10,34),(6,'A música do silêncio',1,'2013-06-15',120,18.65),(7,'Way of Kings',4,'2001-04-20',300,29.53),(8,'Words of Radiance',3,'2014-04-05',200,32.15),(9,'Oathbringer',1,'2016-04-17',150,35.83),(10,'A queda',2,'2018-04-13',50,22.9),(11,'O pistoleiro',2,'2011-04-14',180,40.34),(12,'Mago e vidro',2,'2015-04-12',130,74.43),(13,'A startup enxuta',4,'2007-04-10',140,28.5),(14,'Harry Potter e o cálice de fogo',5,'2004-04-07',60,36.38),(15,'Harry Potter e o enigma do principe',3,'2006-04-30',30,34.99),(16,'Pai rico, pai pobre',7,'1995-04-25',1000,39.99),(17,'As Portas de Pedra',NULL,NULL,NULL,100);

--------------------------------------------------

DROP TABLE IF EXISTS pessoas;
CREATE TABLE pessoas (
  id_pessoa bigint NOT NULL,
  nome varchar(20) DEFAULT NULL,
  sobrenome varchar(45) DEFAULT NULL,
  data_nascimento date DEFAULT NULL,
  residencia varchar(45) DEFAULT NULL,
  funcionario tinyint DEFAULT NULL,
  cpf varchar(11) DEFAULT NULL,
  telefone numeric(11) DEFAULT NULL,
  PRIMARY KEY (id_pessoa),
);

INSERT INTO pessoas VALUES (2,'Lucas','Pereira','1997-04-12','Rua João Baptista Scalco',1,'83432488050',948124864),(3,'Juliana','Silva','1990-03-01','Rua Bromélias',1,'30859454070',949814618),(4,'Paulo','Almeida','1981-07-10','Rua Coronel Aníbal de Andrade',1,'77885049043',977740098),(5,'Sandra','Faria','1997-11-11','Rua Alipio Dutra',1,'98584621075',988828984),(6,'Miguel','Rodrigues','1997-08-01','Rua Carvalho e Melo',1,'92683139012',997324872),(7,'Rafael','Lima','1997-07-26','Praça Frei Paulo',1,'39300155016',998523198),(8,'Vinicius','Souza','1997-06-17','Rua Dorival Ferreira',1,'79715900240',956237008),(9,'Maria','Cavalcanti','1997-03-13','Rua Visconde de São Lourenço',1,'56578773061',947523418),(10,'José','Costa','1997-05-14','Praça Edmundo da Luz Pinto',1,'79002997000',957624235),(11,'Carlos','Oliveira','1997-04-22','Avenida Sargento Carlos Argemiro Camargo',1,'65336136000',952638973),(12,'Valeria','Barbosa','1997-12-02','Rua Julieta',1,'43265967085',923856235),(13,'Adriana','Barros','1999-04-01','Rua Florai',0,'17858811088',956324199),(14,'Pedro','Batista','1992-08-18','Rua Guarajuba',0,'49055201006',NULL),(15,'Ana','Rodrigues','1989-02-11','Travessa São Joaquim',0,'56896350492',997856353),(16,'Vicente','Borges','1997-04-17','Estrada Cruz das Almas',0,'84655966033',998065263),(17,'Antonio','Dias','1977-06-30','Rua Soldado João Rechocoski',0,'68434195003',989645007),(18,'Juliana','Cardoso','1965-10-31','Praça General Santandes',0,'66101712036',996235986),(19,'João','Freitas','1982-12-12','Rua São Fileias',0,'67052297051',NULL),(20,'Francisco','Martins','1986-11-09','Rua Mestre Vitalino',0,'1906371024',990970777),(21,'Camila','Campos','1995-07-04','Rua Herval Rossano',0,'83230880048',NULL),(22,'Bruna','Machado','1987-09-07','Rua Mestre Vitalino',0,'64610073005',909872347),(23,'José','Casta','1983-02-10','Rua Herval Rossano',0,'70859369080',958726393),(24,'Luiz','Gonçalves','1980-02-01','Rua Magistrado',0,'14882069024',987792534),(25,'Mateus','Castro','1978-01-09','Rua Ana Barbosa',0,'11102184004',951478624),(26,'Guilherme','Nunes','1994-06-27','Rua Narcelio de Queiros',0,'50613587081',998236237),(27,'Amanda','Vieira','1993-09-15','Rua Doutor Padilha',0,'1130248003',959263963),(28,'Jessica','Carvalho','1988-11-16','Rua Guidoval',0,'67404832055',983277340);

--------------------------------------------------

DROP TABLE IF EXISTS pedidos;
CREATE TABLE pedidos (
  id_pedidos bigint NOT NULL,
  num_pedido NUMERIC(11) DEFAULT NULL,
  id_livro bigint NOT NULL,
  id_pessoa bigint NOT NULL,
  qtd_pedida NUMERIC(11) DEFAULT NULL,
  PRIMARY KEY (id_pedidos),
);


INSERT INTO pedidos VALUES (3,15,16,20,5),(4,1,6,3,1),(5,2,16,3,12),(6,3,13,3,11),(7,4,10,3,10),(8,5,15,5,2),(9,6,5,9,5),(10,7,11,11,3),(11,8,13,14,2),(12,9,7,16,7),(13,10,8,16,3),(14,11,9,16,2),(15,12,9,17,1),(16,13,3,18,1),(17,14,4,18,1),(18,16,16,20,5);



Atividade proposta:
--01- Retornar da tabela pessoas somente o Miguel
SELECT NOME FROM PESSOAS
WHERE NOME = 'MIGUEL'

--02- Retornar todos os campos da tabela livros o livro cujo valor é de R$40,34
SELECT *  FROM LIVROS
WHERE PRECO = 40.34 
select * FROM livros where round(preco, 2) = 40.34;

--03- Retornar todos os livros cuja data de expedição (dt_expedicao) é do ano de 2011
SELECT * FROM LIVROS 
WHERE DT_EXPEDICAO > '2011-01-01' (ERRADO)
select * from livros Where dt_expedicao LIKE '%2011%'

--Forma mais avançada
select * from livros Where YEAR(dt_expedicao) = '2011'
SELECT * from livros where DATEPART ( YEAR , dt_expedicao ) = '2011';

--04- Retornar todos os livros cuja data de expedição (dt_expedicao) é do mês de novembro
SELECT * FROM LIVROS 
WHERE dt_expedicao BETWEEN '2011-11-01' AND '2011-11-30'

select * from livros Where dt_expedicao LIKE '%-11-%'

--Forma mais avançada
select * from livros Where MONTH(dt_expedicao) = '11'
SELECT * from livros where DATEPART ( MONTH , dt_expedicao ) = '11';

--05- Retornar todos os livros cuja data de expedição (dt_expedicao) está entre o intervalo do dia 5 e 20 de qualquer mês e ano.
SELECT * FROM LIVROS
WHERE DAY(dt_expedicao) BETWEEN 5 AND 20;
-- between retorna o intervalo fechado


--06- Trazer da tabela livros todos os campos e livros cuja edição não é nula.
SELECT * FROM LIVROS 
WHERE EDICAO IS NOT NULL

--07- Trazer da tabela livros todos os livros com seus respectivos campos cuja edição é igual ou maior do que 2 e a quantidade no estoque (qtd_estoque) é menor do que 100
SELECT * FROM LIVROS 
WHERE EDICAO >= 2 AND QTD_ESTOQUE < 100

--08- Trazer todos os livros com seus respectivos campos cuja edição é maior do que 5 e a quantidade no estoque (qtd_estoque) é maior do que 100
SELECT * FROM LIVROS 
WHERE EDICAO > 5 AND QTD_ESTOQUE > 100


--09- Retornar todos os pedidos cuja quantidade pedida (qtd_pedida) NÃO é menor do que 5
SELECT * FROM PEDIDOS 
WHERE NOT QTD_PEDIDA  <=  5 (errado)

SELECT * FROM PEDIDOS 
WHERE NOT QTD_PEDIDA  < 5

select * from pedidos Where qtd_pedida >=5 


--10- Retornar todos os livros que possui Harry no nome
SELECT * FROM lIVROS
WHERE NOME LIKE '%HARRY%'

--11- Retornar todos os livros que possui silêncio no nome (cuidado com o caractere acentuado)
SELECT * FROM lIVROS
WHERE NOME LIKE '%silêncio%'

SELECT * from livros WHERE nome LIKE '%sil%ncio';

--12- Retornar todos os livros com ao menos um caractere
SELECT * FROM LIVROS
WHERE LEN(NOME) > 0;

select * from livros Where nome Like '%_'
SELECT * FROM livros WHERE nome LIKE '_%';


--13- Retornar todos os livros com exatamente 12 caracteres
SELECT * FROM LIVROS
WHERE LEN(NOME) = 12;

select * from livros Where nome Like '____________'

--14- Retorna todos os livros com ao menos um caractere
SELECT * FROM LIVROS
WHERE LEN(NOME) > 0;

select * from livros Where nome Like '%_'
SELECT * FROM livros WHERE nome LIKE '_%';


--15- Retorna todos os livros onde a qtd_estoque é exatamente igual a 100 ou 120 ou 200 ou 60
SELECT * FROM lIVROS
WHERE QTD_ESTOQUE > 60 OR QTD_ESTOQUE > 100 OR QTD_ESTOQUE > 120 OR QTD_ESTOQUE > 200 (ERRADO)

SELECT * FROM livros WHERE qtd_estoque = 100 OR qtd_estoque = 120 OR qtd_estoque = 200 OR qtd_estoque = 60;

--16- Retornar todas as pessoas que nasceram (data_nascimento) entre 1950 e 2010
SELECT * FROM PESSOAS 
WHERE DATA_NASCIMENTO BETWEEN '1950' AND '2010'


SELECT * FROM pessoas WHERE data_nascimento BETWEEN '1950-01-01' AND '2010-12-31';

--17- Retornar todas as pessoas que nasceram entre 01 de janeiro de 1950 e 31 de dezembro de 2010
SELECT * FROM PESSOAS 
WHERE DATA_NASCIMENTO BETWEEN '1950/01/01' AND '2010/12/31'

select * from pessoas Where data_nascimento Between '1950-01-01' AND '2010-12-31'

--18- Retornar da tabela pessoa todas o ano (somente o ano) de nascimento de cada uma delas alterando o nome da coluna data_nascimento para "Ano Nascimento"
SELECT YEAR(DATA_NASCIMENTO) AS 'Ano Nascimento' FROM PESSOAS

select nome, YEAR(data_nascimento) as 'Ano Nascimento' from pessoas
SELECT date_format(data_nascimento, '%Y') AS "Ano Nascimento" FROM pessoas;
SELECT DATEPART (YEAR, data_nascimento) AS "Ano Nascimento" FROM pessoas;


--19- Trazer de livros o nome a qtd_estoque e o preço de cada livro com a seguinte fórmula: preço = (5 * preço) com 2 casas decimais.
      --Altere o nome da coluna nome para "Nome", qtd_estoque para "Estoque" e preco para "Preço"
SELECT NOME AS Nome, QTD_ESTOQUE AS Estoque, ROUND(5 * PRECO, 2) AS Preço FROM LIVROS

select nome as 'Nome', qtd_estoque as 'Estoque', round((preco * 5) ,2) as 'Preço' from livros

--20- Busca o nome de todas as pessoas da tabela pessoa retornando se possui o contato.
       --Caso não possua retornar 'Não' e se possuir retornar 'Sim' em uma coluna chamada 'Contato'

SELECT NOME,
	CASE WHEN TELEFONE IS NOT NULL THEN 'Sim' 
	ELSE 'Não' 
	END AS Contato 
FROM PESSOAS;

SELECT nome, ELSE 'NÃO' END AS 'Realizou algum pedido'
FROM PESSOAS INNER JOIN PEDIDOS ON PEDIDOS.ID_PESSOA = PESSOAS.ID_PESSOA
ORDER BY NOME ASC
	CASE
	WHEN telefone IS NULL THEN 'Não'
	ELSE 'Sim' END AS 'Contato'
FROM pessoas;


-- 21- Retornar todos os campos da tabela livros com limite de 5 registros 
SELECT TOP 5 * FROM LIVROS

-- 22- Retornar todas as pessoas que realizaram algum pedido
SELECT PES.ID_PESSOA, PES.NOME, PES.SOBRENOME, PED.ID_PEDIDOS, PED.QTD_PEDIDA AS 'QUANTIDADE PEDIDA' 
FROM PESSOAS PES 
INNER JOIN PEDIDOS PED ON PED.ID_PESSOA = PES.ID_PESSOA
WHERE QTD_PEDIDA > 0
ORDER BY PES.ID_PESSOA ASC


select pessoa.id_pessoa, pessoa.nome, pedido.id_pedidos, pedido.id_pessoa
from pessoas pessoa, pedidos pedido
where pessoa.id_pessoa = pedido.id_pessoa
	
-- 23- Retornar os campos nome e total da tabela livros com o seguinte critério: total = qtd_estoque * preco com 2 casas decimais Utilizar subquery para resolver esse exercício.
-- Obs.: A utilização de subquery não é necessária, está sendo pedido apenas como exercício. - médio
SELECT NOME, TOTAL FROM LIVROS 
LEFT JOIN (SELECT ROUND(QTD_ESTOQUE * PRECO, 2) AS 'TOTAL', ID_LIVRO FROM LIVROS) AS TESTE ON teste.ID_Livro = Livros.Id_livro

select nome, 
(select round(qtd_estoque * preco,2) as 'Total')
from livros

--24- Retornar todas as colunas de livros cujo número de pedidos seja maior ou igual a 5
SELECT * FROM LIVROS L 
INNER JOIN PEDIDOS P ON P.ID_LIVRO = L.ID_LIVRO
WHERE QTD_PEDIDA >= 5
ORDER BY L.NOME ASC

Select Distinct liv.*
From livros_Aten As liv, pedidos_Aten As ped
Where liv.id_livro In (Select ped.id_livro From livros_Aten As liv, pedidos_Aten As ped Group By ped.id_livro Having liv.qtd_estoque < Sum(ped.qtd_pedida)*2);


--25- Retornar todos os livros cuja a quantidade de estoque é menor do que 2 vezes a quantidade pedida 
SELECT * FROM LIVROS L 
INNER JOIN PEDIDOS P ON L.ID_LIVRO = P.ID_LIVRO
WHERE QTD_ESTOQUE < 2 * QTD_PEDIDA;

SELECT * FROM LIVROS L 
INNER JOIN PEDIDOS P ON L.ID_LIVRO = P.ID_LIVRO
WHERE QTD_ESTOQUE <= 2 * QTD_PEDIDA;


SELECT * FROM livros WHERE qtd_estoque <= ANY (SELECT (2 * qtd_pedida) FROM pedidos);


--26- Unir as tabelas livros e pedidos (à esquerda) e trazer todos os campos.
SELECT * 
FROM LIVROS L 
LEFT JOIN PEDIDOS P ON L.ID_LIVRO = P.ID_LIVRO

--27- Juntar as tabelas livros e pedidos e retornar os livros já comprados.
SELECT L.NOME,
CASE 
	WHEN NUM_PEDIDO > 0 THEN 'SIM' 
	ELSE 'NÃO' 
	END AS 'COMPRADO'
FROM LIVROS L 
INNER JOIN PEDIDOS P ON P.ID_LIVRO = L.ID_LIVRO


select * from livros livro, pedidos pedido
where livro.id_livro = pedido.id_livro and pedido.qtd_pedida >0

--28- Retornar o nome dos livros e a quantidade total já pedida
SELECT L.NOME, P.QTD_PEDIDA AS "TOTAL PEDIDO",
CASE WHEN NUM_PEDIDO > 0 THEN 'SIM' ELSE 'NÃO' END AS 'COMPRADO'
FROM LIVROS L INNER JOIN PEDIDOS P ON P.ID_LIVRO = L.ID_LIVRO
ORDER BY P.QTD_PEDIDA ASC


SELECT l.nome, SUM(p.qtd_pedida)  AS 'Total Pedido'
FROM livros l
LEFT JOIN pedidos p ON l.id_livro = p.id_livro
GROUP BY l.nome;

--29- Juntar a tabela pessoa com a de pedidos
SELECT DISTINCT * 
FROM PESSOAS P 
INNER JOIN PEDIDOS PE ON P.ID_PESSOA = P.ID_PESSOA

SELECT pessoa.*, pedido.*
FROM pessoas pessoa 
FULL OUTER JOIN pedidos pedido
on pessoa.id_pessoa = pedido.id_pessoa

--30- Utilizar a tebela pedidos à direita e retornar a intercessão com a tabela de pessoas
SELECT DISTINCT *
FROM PEDIDOS PED
RIGHT JOIN PESSOAS PES ON PES.ID_PESSOA = PED.ID_PESSOA;

SELECT ps.nome, ps.sobrenome, ps.telefone,
       pd.num_pedido, pd.id_livro
FROM pessoas ps
RIGHT JOIN pedidos pd ON pd.id_pessoa = ps.id_pessoa;

--31- Unir as tabelas pedidos e pessoas e retornar os elementos comuns 
SELECT * 
FROM PEDIDOS PE 
INNER JOIN PESSOAS PA ON PE.ID_PESSOA = PA.ID_PESSOA 

SELECT pedido.*,pessoa.*
FROM pedidos pedido
INNER JOIN pessoas pessoa
on pedido.id_pessoa = pessoa.id_pessoa

--32- Criar a relação de todos os elementos da tabela pessoas com a de pedidos (produto cartesiano entre as tabelas - CROSS JOIN)
SELECT * 
FROM PESSOAS AS PES 
CROSS JOIN PEDIDOS AS PED

/*33- Unir o resultado de 2 selects em uma única tabela
    Primeiro) nome e dt_expedicao de livros cujo preco e maior do que 50
    Segundo)  nome e dt_expedicao de livros cujo preco e menor do que 508*/
	
SELECT NOME, PRECO, DT_EXPEDICAO 
FROM LIVROS 
WHERE PRECO > 50 

UNION

SELECT NOME, PRECO, DT_EXPEDICAO 
FROM LIVROS 
WHERE PRECO < 508;	
	
/*34- Unir o resultado de 2 selects em uma única tabela - ?
    Primeiro) nome e dt_expedicao de livros cujo preco e maior do que 50
    Segundo)  nome e dt_expedicao de livros cujo preco e menor do que 100*/

SELECT NOME, PRECO, DT_EXPEDICAO 
FROM LIVROS 
WHERE PRECO > 50 

UNION

SELECT NOME, PRECO, DT_EXPEDICAO 
FROM LIVROS 
WHERE PRECO < 100;

--35- Realizar o FULL ALTER JOIN entre as tabelas livros e pedidos retornando nome, dt_expedicao, num_pedido e id_pessoa
SELECT NOME, DT_EXPEDICAO, NUM_PEDIDO, ID_PESSOA
FROM LIVROS
FULL OUTER JOIN PEDIDOS ON LIVROS.ID_LIVRO = PEDIDOS.ID_LIVRO

SELECT l.nome, l.dt_expedicao,
       p.num_pedido, p.id_pessoa
FROM livros l
LEFT JOIN pedidos p ON p.id_livro = l.id_livro
UNION
SELECT l.nome, l.dt_expedicao,
       p.num_pedido, p.id_pessoa
FROM livros l
RIGHT JOIN pedidos p ON p.id_livro = l.id_livro;

--36- Inserir o livro "o poder das frutas" cuja edição é 1 e a data de expedição é 14/03/2019 com a quantidade em estoque igual a 50 e o preco de R$120,25
INSERT INTO LIVROS (ID_LIVRO, NOME, EDICAO, DT_EXPEDICAO, QTD_ESTOQUE, PRECO)
VALUES (18, 'O poder das frutas', 1, '14-03-2019', 50, '120.25') 

--37- Alterar o nome do livro "o poder das frutas" para "o poder das hortaliças"
UPDATE LIVROS 
SET NOME = 'o poder das hortaliças'
WHERE ID_LIVRO = 18

--38- Excluir o livro "o poder das hortaliças"
DELETE FROM LIVROS 
WHERE ID_LIVRO = 18
