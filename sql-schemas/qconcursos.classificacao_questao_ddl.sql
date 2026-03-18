-- qconcursos.classificacao_questao definition

CREATE TABLE `classificacao_questao` (
  `id_questao` int NOT NULL,
  `id_topico` bigint NOT NULL,
  PRIMARY KEY (`id_questao`,`id_topico`),
  KEY `FKo6qn0ru2hd15uf5tymb726yn7` (`id_topico`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;