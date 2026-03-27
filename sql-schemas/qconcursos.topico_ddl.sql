-- qconcursos.topico definition

CREATE TABLE `topico` (
  `id` bigint NOT NULL AUTO_INCREMENT,
  `nome` varchar(255) NOT NULL,
  `id_pai` bigint DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `FK4s6ve3g45lrj367w9ox5wxr5h` (`id_pai`)
) ENGINE=MyISAM AUTO_INCREMENT=12221 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;