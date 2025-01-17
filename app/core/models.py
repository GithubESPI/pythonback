
from sqlalchemy import Column, String, Integer, Boolean, DateTime, ForeignKey, LargeBinary
from sqlalchemy.orm import relationship
from sqlalchemy.ext.declarative import declarative_base
import datetime

Base = declarative_base()

class User(Base):
    __tablename__ = "User"
    id = Column(String, primary_key=True, unique=True)
    name = Column(String, nullable=True)
    email = Column(String, unique=True, nullable=False)
    emailVerified = Column(DateTime, nullable=True)
    image = Column(String, nullable=True)
    createdAt = Column(DateTime, default=datetime.datetime.utcnow)
    updatedAt = Column(DateTime, onupdate=datetime.datetime.utcnow)

    configurations = relationship("Configuration", back_populates="user")
    generated_excels = relationship("GeneratedExcel", back_populates="user")


class Configuration(Base):
    __tablename__ = "Configuration"
    id = Column(String, primary_key=True, unique=True)
    fileName = Column(String, nullable=False)
    excelUrl = Column(String, nullable=False)
    wordUrl = Column(String, nullable=False)
    userId = Column(String, ForeignKey("User.id"), nullable=False)
    generatedExcel = Column(LargeBinary, nullable=True)
    generatedBulletins = Column(LargeBinary, nullable=True)
    createdAt = Column(DateTime, default=datetime.datetime.utcnow)
    updatedAt = Column(DateTime, onupdate=datetime.datetime.utcnow)

    user = relationship("User", back_populates="configurations")


class GeneratedFile(Base):
    __tablename__ = "GeneratedFile"
    id = Column(Integer, primary_key=True, autoincrement=True)
    filename = Column(String, nullable=False)
    fileType = Column(String, nullable=False)
    fileData = Column(LargeBinary, nullable=False)
    isTemplate = Column(Boolean, default=False)
    templateType = Column(String, nullable=True)
    category = Column(String, nullable=True)
    createdAt = Column(DateTime, default=datetime.datetime.utcnow)
    updatedAt = Column(DateTime, onupdate=datetime.datetime.utcnow)


class GeneratedExcel(Base):
    __tablename__ = "GeneratedExcel"
    id = Column(Integer, primary_key=True, autoincrement=True)
    userId = Column(String, ForeignKey("User.id"), nullable=False)
    templateId = Column(Integer, ForeignKey("GeneratedFile.id"), nullable=False)
    data = Column(LargeBinary, nullable=False)
    createdAt = Column(DateTime, default=datetime.datetime.utcnow)
    updatedAt = Column(DateTime, onupdate=datetime.datetime.utcnow)

    user = relationship("User", back_populates="generated_excels")
    template = relationship("GeneratedFile")
