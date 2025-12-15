--CREATE DATABASE Lab;

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


