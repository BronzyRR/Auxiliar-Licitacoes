from dataclasses import dataclass

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from ArquivosWord import *
import shutil


base = Document()


#DADOS FIXOS
empresa = "MULTIVENDAS COMÉRCIO E SERVIÇOS LTDA"
cnpj = "07.538.900/0001-36"
enderecoSede = "Rua Cecília Brasil, 1274, Centro"

nomeResponsavelTecnico = "JOSEILDO SOARES DE SOUZA"
registroCREA = "918.444.926"
cpf_responsavel_tecnico = "329.171.202-15"

representante_legal = "JOSEILDO SOARES DE SOUZA"
cargo_representante_legal = "SÓCIO-GERENTE"
rg = "1973885"
orgaoEmissor = "SSP-PA"
cpf = "329.171.202-15"


#DADOS QUE VARIAM DE UMA LICITAÇÃO PARA OUTRA
concorrencia = "CONCORRÊNCIA Nº. 022/2023"

processo = "PROCESSO Nº. 027632/2023 - SEMGES"

objeto = "CONTRATAÇÃO DE EMPRESA ESPECIALIZADA EM OBRAS E SERVIÇOS DE ENGENHARIA, PARA CONSTRUÇÃO DO ABRIGO " \
         "INSTITUCIONAL DE IDOSOS, NO MUNICIPIO DE BOA VISTA-RR."

data = "11 de março de 2024"

horarioAbertura = "09h:00min"

valor_total_proposta = "6.057.254,74"

prazo_validade_proposta = "60 (sessenta) dias"

prazo_garantia = "05 (cinco) anos"

prazo_execucao_contrato = "240 (duzentos e quarenta) dias"



input("JÁ EDITOU OS DADOS DA LICITAÇÃO AQUI NO CÓDIGO?")

enderecoSalvarHabilitacao = input("Informe o endereço em que deseja salvar os documentos da habilitação: ")
enderecoSalvarHabilitacao = enderecoSalvarHabilitacao[1:-1]

enderecoSalvarProposta = input("Informe o endereço em que deseja salvar os documentos da proposta: ")
enderecoSalvarProposta = enderecoSalvarProposta[1:-1]


#DOCUMENTOS DA HABILILTAÇÃO
criar_documento01_aceitacao(concorrencia, processo, nomeResponsavelTecnico, cpf_responsavel_tecnico, registroCREA, data,enderecoSalvarHabilitacao)

criar_documento02_equipe_tecnica(concorrencia, processo, nomeResponsavelTecnico, cpf_responsavel_tecnico, registroCREA,
                                 data, representante_legal, cargo_representante_legal, enderecoSalvarHabilitacao)

criar_documento03_pertence_quadro_empresa(concorrencia, processo, nomeResponsavelTecnico, cpf_responsavel_tecnico,
                                          registroCREA, data, representante_legal, cargo_representante_legal, enderecoSalvarHabilitacao)

criar_documento04_alvara(concorrencia, processo, representante_legal, cargo_representante_legal, objeto,
                         cnpj, enderecoSede, rg, orgaoEmissor, data, enderecoSalvarHabilitacao)

criar_documento05_total_conhecimento(empresa, cnpj, concorrencia, objeto, data, representante_legal,
                                     rg, orgaoEmissor, enderecoSalvarHabilitacao)

criar_documento06_declaracao_diversas(empresa, cnpj, enderecoSede, concorrencia, processo, data,
                                      representante_legal, rg, orgaoEmissor, enderecoSalvarHabilitacao)

criar_documento07_execucao_de_acordo(empresa, cnpj, concorrencia, objeto, data, representante_legal,
                                     rg, orgaoEmissor, enderecoSalvarHabilitacao)

criar_documento08_pleno_conhecimento(empresa, cnpj, concorrencia, objeto, data, representante_legal,
                                     rg, orgaoEmissor, enderecoSalvarHabilitacao)

criar_documento09_recebimento_edital(empresa, cnpj, concorrencia, objeto, data, representante_legal,
                                     rg, orgaoEmissor, enderecoSalvarHabilitacao)

criar_documento10_ensaios_tecnologicos(empresa, cnpj, concorrencia, objeto, data, representante_legal,
                                       rg, orgaoEmissor, enderecoSalvarHabilitacao)

criar_documento11_informacoes_contrato(data, enderecoSalvarHabilitacao)

criar_envelope_propostas(concorrencia, processo, objeto, data, horarioAbertura, enderecoSalvarHabilitacao)

#DOCUMENTOS PARA PROPOSTA
criar_documento_carta_proposta(concorrencia, processo, data, horarioAbertura, objeto, valor_total_proposta,
                               prazo_validade_proposta, prazo_garantia, prazo_execucao_contrato, nomeResponsavelTecnico,
                               rg, orgaoEmissor, enderecoSalvarProposta)

criar_declaracao_independente_proposta(empresa, cnpj, concorrencia, processo, data, representante_legal,
                                       rg, orgaoEmissor, cpf, enderecoSalvarProposta)


# CAPAS E DOCUMENTOS PARA EDITAR MANUALMENTE
shutil.copyfile("C:/Users/willi/PycharmProjects/MultivendasLicitacao/configs/(12) CAPA DOCUMENTO CREDENCIAMENTO.doc",
                "{}/(12) CAPA DOCUMENTO CREDENCIAMENTO.doc".format(enderecoSalvarHabilitacao))

shutil.copyfile("C:/Users/willi/PycharmProjects/MultivendasLicitacao/configs/(13) CAPA DOCUMENTO HABILITAÇÃO.doc",
                "{}/(13) CAPA DOCUMENTO HABILITAÇÃO.doc".format(enderecoSalvarHabilitacao))

shutil.copyfile("C:/Users/willi/PycharmProjects/MultivendasLicitacao/configs/Capa_Proposta_preço.docx",
                "{}/(04) CAPA PROPOSTA DE PREÇOS.docx".format(enderecoSalvarProposta))


criar_capa_cd_proposta(concorrencia, processo, objeto, enderecoSalvarProposta)