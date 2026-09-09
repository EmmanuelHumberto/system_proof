#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Regras de classificação de comprovantes para despesas da planilha."""


def classificar_despesa(cat, c):
    alvo = (
        c.get("arquivo", "")
        + " "
        + (c.get("recebedor") or "")
        + " "
        + (c.get("tipo") or "")
    ).lower()

    if cat == "Alimentação":
        if any(k in alvo for k in ("suplemento", "ultrafarma", "danfe")):
            return "Suplementos e alimentação especial"
        return "Supermercado e alimentação domiciliar"

    if cat == "Comunicação e tecnologia":
        return "Plano de telefonia celular"

    if cat == "Lazer e convivência":
        if any(k in alvo for k in ("netflix", "spoti", "ifood club", "uber one", "streaming")):
            return "Assinaturas de streaming – quota-parte"
        return "Cinema, parques e passeios"

    if cat == "Moradia":
        if "condomínio" in c.get("tipo", "").lower() or "pacto" in alvo:
            return "Condomínio – quota-parte"
        if "internet" in alvo:
            return "Internet residencial – quota-parte"
        if "energia" in c.get("tipo", "").lower():
            return "Energia elétrica – quota-parte"
        if "água" in c.get("tipo", "").lower():
            return "Água e esgoto – quota-parte"
        if "gás" in c.get("tipo", "").lower():
            return "Gás – quota-parte"
        if "iptu" in alvo or "pref mun" in alvo or "prefeitura" in alvo:
            return "IPTU – quota-parte"
        if "maria efigênia" in alvo or "faxina" in alvo or "diarista" in alvo:
            return "Outras despesas comprovadas"
        return "Aluguel ou financiamento – quota-parte do adolescente"

    if cat == "Saúde":
        if "nfs-e" in c.get("tipo", "").lower() or "odonto" in alvo:
            return "Tratamento odontológico"
        if "receita" in c.get("tipo", "").lower():
            return "Consultas médicas particulares"
        if any(k in alvo for k in ("otica", "ótica", "lente")):
            return "Óculos, lentes e itens ortopédicos"
        if any(k in alvo for k in ("raia", "drogasil", "drogaria", "araujo", "cartao de todos", "cartão de todos")):
            return "Medicamentos contínuos"
        return "Consultas médicas particulares"

    if cat == "Transporte":
        return "Aplicativos de transporte"

    if cat == "Educação":
        if "escola" in c.get("tipo", "").lower() or "target" in alvo:
            return "Mensalidade escolar"
        if any(k in alvo for k in ("refor", "aula", "emmanuel")):
            return "Curso preparatório / profissionalizante"
        if "enem" in alvo or "inep" in alvo:
            return "Curso preparatório / profissionalizante"
        if "livro" in c.get("tipo", "").lower() or "musicalia" in alvo:
            return "Livros didáticos e paradidáticos"
        if "música" in alvo or "violino" in alvo or "camila" in alvo or "victor" in alvo or "vinicius" in alvo:
            return "Cursos de música, arte ou cultura"
        return "Mensalidade escolar"

    if cat == "Vestuário e higiene":
        if any(k in alvo for k in ("riachuelo", "renner")):
            return "Roupas"
        if any(k in alvo for k in ("studio mirra", "adilson", "cabeleireiro", "barbearia")):
            return "Cabeleireiro / cuidados pessoais"
        return "Roupas"

    return "Outras despesas comprovadas"
