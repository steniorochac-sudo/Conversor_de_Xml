from decimal import Decimal
from datetime import datetime
from typing import List, Optional, Any
from pydantic import BaseModel, Field, ConfigDict, field_validator
from fiscal_workflow.models.models import RegimeTributario, StatusApuracao

# ==========================================
# SCHEMAS PARA AJUSTE DOCUMENTO (AUDITORIA)
# ==========================================

class AjusteCreate(BaseModel):
    valor_total_ajuste: Decimal = Field(
        ..., 
        description="Valor a somar/subtrair do total (positivo ou negativo)",
        decimal_places=2
    )
    justificativa: str = Field(
        ..., 
        min_length=5, 
        max_length=255, 
        description="Justificativa legal ou gerencial do ajuste"
    )
    usuario: Optional[str] = Field("usuario_sistema", max_length=100)

class AjusteResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    documento_id: int
    valor_total_ajuste: Decimal
    justificativa: str
    usuario: str
    data_ajuste: datetime

# ==========================================
# SCHEMAS PARA EMPRESA
# ==========================================

class EmpresaCreate(BaseModel):
    cnpj: str = Field(..., min_length=14, max_length=14, description="CNPJ composto apenas de 14 dígitos")
    razao_social: str = Field(..., min_length=2, max_length=255)
    regime_tributario: RegimeTributario = Field(..., description="Regime tributário da empresa")
    rbt12: Optional[Decimal] = Field(Decimal("0.00"), description="Faturamento acumulado dos últimos 12 meses")
    folha12: Optional[Decimal] = Field(Decimal("0.00"), description="Folha de salários dos últimos 12 meses")
    sujeito_fator_r: Optional[bool] = Field(False, description="Atividade sujeita a Fator R")
    categoria_simples: Optional[str] = Field("Serviços (Anexo III)", description="Atividade / Anexo do Simples Nacional")
    cnae: Optional[str] = Field(None, description="CNAE da empresa")

    @field_validator('cnpj')
    @classmethod
    def validar_cnpj_numerico(cls, v: str) -> str:
        if not v.isdigit():
            raise ValueError("O CNPJ deve conter apenas dígitos numéricos.")
        return v

class EmpresaResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    cnpj: str
    razao_social: str
    regime_tributario: RegimeTributario
    rbt12: Decimal
    folha12: Decimal
    sujeito_fator_r: bool
    categoria_simples: str
    cnae: Optional[str] = None

class EmpresaUpdate(BaseModel):
    razao_social: str = Field(..., min_length=2, max_length=255)
    regime_tributario: RegimeTributario = Field(..., description="Regime tributário da empresa")
    rbt12: Decimal = Field(Decimal("0.00"), description="Faturamento acumulado dos últimos 12 meses")
    folha12: Decimal = Field(Decimal("0.00"), description="Folha de salários dos últimos 12 meses")
    sujeito_fator_r: bool = Field(False, description="Atividade sujeita a Fator R")
    categoria_simples: str = Field("Serviços (Anexo III)", description="Atividade / Anexo do Simples Nacional")
    cnae: Optional[str] = Field(None, description="CNAE da empresa")

# ==========================================
# SCHEMAS PARA DOCUMENTO FISCAL
# ==========================================

class DocumentoResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    empresa_id: int
    chave_acesso: str
    tipo_documento: str
    tipo_operacao: str
    valor_total: Decimal
    status_apuracao: StatusApuracao
    cstat: str
    itens: Optional[List[Any]] = None
    data_emissao: Optional[datetime] = None
    
    # Auditoria de Ajustes Manuais
    ajustes: List[AjusteResponse] = []
    
    # Campo calculado dinamicamente pelo modelo
    valor_final: Decimal
