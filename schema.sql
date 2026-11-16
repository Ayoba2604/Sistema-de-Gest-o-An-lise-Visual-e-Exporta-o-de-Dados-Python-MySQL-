CREATE DATABASE IF NOT EXISTS estoque_gestao
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;

USE estoque_gestao;

CREATE TABLE IF NOT EXISTS produtos (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nome VARCHAR(100) NOT NULL,
    categoria VARCHAR(60) NOT NULL,
    preco_unitario DECIMAL(10, 2) NOT NULL,
    estoque_inicial INT NOT NULL DEFAULT 0,
    criado_em TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CHECK (preco_unitario >= 0),
    CHECK (estoque_inicial >= 0)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS movimentacoes (
    id INT AUTO_INCREMENT PRIMARY KEY,
    produto_id INT NOT NULL,
    tipo ENUM('Entrada', 'Saida') NOT NULL,
    quantidade INT NOT NULL,
    data_movimento DATE NOT NULL,
    valor_unitario DECIMAL(10, 2) NOT NULL,
    observacao VARCHAR(255) NOT NULL DEFAULT '',
    CONSTRAINT fk_movimentacoes_produtos
        FOREIGN KEY (produto_id) REFERENCES produtos(id)
        ON UPDATE CASCADE
        ON DELETE RESTRICT,
    CHECK (quantidade > 0),
    CHECK (valor_unitario >= 0),
    INDEX idx_movimentacoes_data (data_movimento),
    INDEX idx_movimentacoes_tipo (tipo)
) ENGINE=InnoDB;

INSERT INTO produtos (nome, categoria, preco_unitario, estoque_inicial)
SELECT dados.nome, dados.categoria, dados.preco_unitario, dados.estoque_inicial
FROM (
    SELECT 'Caderno universitario' AS nome, 'Papelaria' AS categoria, 24.90 AS preco_unitario, 20 AS estoque_inicial
    UNION ALL SELECT 'Caneta azul', 'Papelaria', 2.50, 50
    UNION ALL SELECT 'Mochila escolar', 'Acessorios', 139.90, 8
    UNION ALL SELECT 'Garrafa termica', 'Acessorios', 59.90, 12
    UNION ALL SELECT 'Calculadora', 'Eletronicos', 49.90, 10
) AS dados
WHERE NOT EXISTS (SELECT 1 FROM produtos);

INSERT INTO movimentacoes
    (produto_id, tipo, quantidade, data_movimento, valor_unitario, observacao)
SELECT p.id, dados.tipo, dados.quantidade, dados.data_movimento, dados.valor_unitario, dados.observacao
FROM (
    SELECT 'Caderno universitario' AS produto, 'Entrada' AS tipo, 12 AS quantidade, '2026-01-08' AS data_movimento, 24.90 AS valor_unitario, 'Reposicao de janeiro' AS observacao
    UNION ALL SELECT 'Caneta azul', 'Saida', 8, '2026-01-12', 2.50, 'Venda'
    UNION ALL SELECT 'Mochila escolar', 'Entrada', 5, '2026-01-20', 139.90, 'Compra do fornecedor'
    UNION ALL SELECT 'Garrafa termica', 'Saida', 2, '2026-02-02', 59.90, 'Venda'
    UNION ALL SELECT 'Calculadora', 'Entrada', 4, '2026-02-09', 49.90, 'Reposicao'
    UNION ALL SELECT 'Caderno universitario', 'Saida', 6, '2026-02-18', 24.90, 'Venda'
    UNION ALL SELECT 'Caneta azul', 'Entrada', 20, '2026-03-03', 2.50, 'Reposicao'
    UNION ALL SELECT 'Mochila escolar', 'Saida', 3, '2026-03-14', 139.90, 'Venda'
    UNION ALL SELECT 'Garrafa termica', 'Entrada', 6, '2026-04-04', 59.90, 'Compra do fornecedor'
    UNION ALL SELECT 'Calculadora', 'Saida', 2, '2026-04-19', 49.90, 'Venda'
    UNION ALL SELECT 'Caderno universitario', 'Entrada', 10, '2026-05-07', 24.90, 'Reposicao'
    UNION ALL SELECT 'Caneta azul', 'Saida', 15, '2026-05-21', 2.50, 'Venda'
    UNION ALL SELECT 'Mochila escolar', 'Entrada', 4, '2026-06-02', 139.90, 'Compra do fornecedor'
    UNION ALL SELECT 'Garrafa termica', 'Saida', 3, '2026-06-17', 59.90, 'Venda'
    UNION ALL SELECT 'Calculadora', 'Entrada', 7, '2026-07-11', 49.90, 'Reposicao'
) AS dados
JOIN produtos AS p ON p.nome = dados.produto
WHERE NOT EXISTS (SELECT 1 FROM movimentacoes);
