from enum import Enum
from decimal import Decimal
from datetime import datetime
from typing import List, Optional
from sqlalchemy import String, ForeignKey, Numeric, JSON, DateTime, func, Enum as SQLEnum
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship

class Base(DeclarativeBase):
    """Classe base declarativa do SQLAlchemy para todos os modelos."""
    pass

class RegimeTributario(str, Enum):
    """Enum para os regimes tributários suportados, garantindo tipagem estrita."""
    SIMPLES_NACIONAL = "Simples Nacional"
    LUCRO_PRESUMIDO = "Lucro Presumido"
    LUCRO_REAL = "Lucro Real"

class StatusApuracao(str, Enum):
    """Enum para o status de apuração de um documento fiscal na Staging Area."""
    PENDENTE = "Pendente"
    EM_REVISAO = "Em Revisão"
    ENCERRADO = "Encerrado"

class Empresa(Base):
    """Modelo representando uma Empresa cadastrada no sistema."""
    __tablename__ = "empresas"

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    cnpj: Mapped[str] = mapped_column(String(14), unique=True, index=True, nullable=False)
    razao_social: Mapped[str] = mapped_column(String(255), nullable=False)
    regime_tributario: Mapped[RegimeTributario] = mapped_column(
        SQLEnum(RegimeTributario, native_enum=False), 
        nullable=False
    )
    rbt12: Mapped[Decimal] = mapped_column(Numeric(15, 2), default=Decimal("0.00"), nullable=False)
    folha12: Mapped[Decimal] = mapped_column(Numeric(15, 2), default=Decimal("0.00"), nullable=False)
    sujeito_fator_r: Mapped[bool] = mapped_column(default=False, nullable=False)
    categoria_simples: Mapped[str] = mapped_column(String(50), default="Serviços (Anexo III)", nullable=False)
    cnae: Mapped[Optional[str]] = mapped_column(String(7), nullable=True)

    # Relacionamento de um para muitos com DocumentoFiscal
    documentos: Mapped[List["DocumentoFiscal"]] = relationship(
        "DocumentoFiscal", 
        back_populates="empresa", 
        cascade="all, delete-orphan"
    )

    def __repr__(self) -> str:
        return f"<Empresa id={self.id} cnpj={self.cnpj} razao_social={self.razao_social}>"

class DocumentoFiscal(Base):
    """Modelo representando uma Nota Fiscal (NF-e/NFC-e) associada a uma empresa."""
    __tablename__ = "documentos_fiscais"

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    empresa_id: Mapped[int] = mapped_column(
        ForeignKey("empresas.id", ondelete="CASCADE"), 
        index=True, 
        nullable=False
    )
    chave_acesso: Mapped[str] = mapped_column(String(60), unique=True, index=True, nullable=False)
    tipo_documento: Mapped[str] = mapped_column(String(20), nullable=False)
    tipo_operacao: Mapped[str] = mapped_column(String(20), default="Saída", nullable=False)
    valor_total: Mapped[Decimal] = mapped_column(Numeric(15, 2), nullable=False)
    status_apuracao: Mapped[StatusApuracao] = mapped_column(
        SQLEnum(StatusApuracao, native_enum=False), 
        default=StatusApuracao.PENDENTE, 
        nullable=False
    )
    itens: Mapped[list] = mapped_column(JSON, nullable=True)
    cstat: Mapped[str] = mapped_column(String(10), default="100", nullable=False)
    data_emissao: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)

    # Relacionamento reverso com Empresa
    empresa: Mapped["Empresa"] = relationship("Empresa", back_populates="documentos")

    # Relacionamento de um para muitos com os ajustes de auditoria
    ajustes: Mapped[List["AjusteDocumento"]] = relationship(
        "AjusteDocumento", 
        back_populates="documento", 
        cascade="all, delete-orphan"
    )

    @property
    def valor_final(self) -> Decimal:
        """Retorna o valor final calculado: valor original + soma de todos os ajustes registrados."""
        soma_ajustes = sum((ajuste.valor_total_ajuste for ajuste in self.ajustes), Decimal("0.00"))
        return self.valor_total + soma_ajustes

    def __repr__(self) -> str:
        return f"<DocumentoFiscal id={self.id} chave_acesso={self.chave_acesso} valor={self.valor_total} status={self.status_apuracao}>"

class AjusteDocumento(Base):
    """Modelo de Auditoria contendo alterações manuais e justificativas feitas pelo usuário."""
    __tablename__ = "ajustes_documentos"

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    documento_id: Mapped[int] = mapped_column(
        ForeignKey("documentos_fiscais.id", ondelete="CASCADE"), 
        index=True, 
        nullable=False
    )
    valor_total_ajuste: Mapped[Decimal] = mapped_column(
        Numeric(15, 2), 
        default=Decimal("0.00"), 
        nullable=False
    )
    justificativa: Mapped[str] = mapped_column(String(255), nullable=False)
    usuario: Mapped[str] = mapped_column(String(100), default="usuario_sistema", nullable=False)
    data_ajuste: Mapped[datetime] = mapped_column(DateTime, default=func.now(), nullable=False)

    # Relacionamento reverso com DocumentoFiscal
    documento: Mapped["DocumentoFiscal"] = relationship("DocumentoFiscal", back_populates="ajustes")

    def __repr__(self) -> str:
        return f"<AjusteDocumento id={self.id} doc_id={self.documento_id} ajuste={self.valor_total_ajuste} usuario={self.usuario}>"
