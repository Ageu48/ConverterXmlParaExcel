import xmltodict
import os
import pandas as pd


def pegar_infos(nome_arquivo, valores):
    print(f"Copiado → {nome_arquivo}")
    with open(f'NFE/{nome_arquivo}', "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)

        if "NFe" in dic_arquivo:
            infos_nf = dic_arquivo["NFe"]["infNFe"]
        else:
            infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]

        # Dados campo identificador(ide)
        cuf = infos_nf["ide"]["cUF"]
        cnf = infos_nf["ide"]["cNF"]
        natop = infos_nf["ide"]["natOp"]
        mod = infos_nf["ide"]["mod"]
        serie = infos_nf["ide"]["serie"]
        nNF = infos_nf["ide"]["nNF"]
        dhEmi = infos_nf["ide"]["dhEmi"]
        tpNF = infos_nf["ide"]["tpNF"]
        idDest = infos_nf["ide"]["idDest"]
        cMunFG = infos_nf["ide"]["cMunFG"]
        tpImp = infos_nf["ide"]["tpImp"]
        tpEmis = infos_nf["ide"]["tpEmis"]
        cDV = infos_nf["ide"]["cDV"]
        tpAmb = infos_nf["ide"]["tpAmb"]
        finNFe = infos_nf["ide"]["finNFe"]
        indFinal = infos_nf["ide"]["indFinal"]
        indPres = infos_nf["ide"]["indPres"]
        procEmi = infos_nf["ide"]["procEmi"]
        verProc = infos_nf["ide"]["verProc"]

        # Dados campo emitente(emit)
        cnpj = infos_nf["emit"]["CNPJ"]
        xNome = infos_nf["emit"]["xNome"]
        xFant = infos_nf["emit"]["xFant"]
        Emit_xLgr = infos_nf["emit"]["enderEmit"]["xLgr"]
        Emit_nro = infos_nf["emit"]["enderEmit"]["nro"]
        Emit_xBairro = infos_nf["emit"]["enderEmit"]["xBairro"]
        Emit_cMun = infos_nf["emit"]["enderEmit"]["cMun"]
        Emit_xMun = infos_nf["emit"]["enderEmit"]["xMun"]
        Emit_UF = infos_nf["emit"]["enderEmit"]["UF"]
        Emit_CEP = infos_nf["emit"]["enderEmit"]["CEP"]
        Emit_cPais = infos_nf["emit"]["enderEmit"]["cPais"]
        Emit_xPais = infos_nf["emit"]["enderEmit"]["xPais"]
        Emit_fone = infos_nf["emit"]["enderEmit"]["fone"]
        ie = infos_nf["emit"]["IE"]
        crt = infos_nf["emit"]["CRT"]
        valores.append([cuf, cnf, natop, mod, serie, nNF, dhEmi, tpNF, idDest, cMunFG, tpImp, tpEmis, cDV, tpAmb,
                        finNFe, indFinal, indPres, procEmi, verProc, cnpj, xNome, xFant, Emit_xLgr,
                        Emit_nro, Emit_xBairro, Emit_cMun, Emit_xMun, Emit_UF,
                        Emit_CEP, Emit_cPais, Emit_xPais, Emit_fone])


lista_arquivos = os.listdir("NFE")

colunas = ["ide_cuf", "ide_cnf", "nat_op", "mod", "serie", "nNF", "dhEmi", "tpNF", "idDest", "cMunFG", "tpImp",
           "tpEmis", "cDV", "tpAmb", "finNFe", "indFinal", "indPres", "procEmi", "verProc", "cnpj", "xNome", "xFant",
           "enderEmit_xLgr", "enderEmit_nro", "enderEmit_xBairro", "enderEmit_cMun", "enderEmit_xMun", "enderEmit_UF",
           "enderEmit_CEP", "enderEmit_cPais", "enderEmit_xPais", "enderEmit_fone"
]
valores = []
for arquivo in lista_arquivos:
    pegar_infos(arquivo, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel("notas_fiscais.xlsx", index=False)
