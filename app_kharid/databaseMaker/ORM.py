from sqlalchemy import create_engine, Column, Integer, String
from sqlalchemy.orm import sessionmaker
from sqlalchemy.orm import declarative_base

db_url = 'sqlite:///mydatabase.db'
engine = create_engine(db_url)
Base = declarative_base()

class Company(Base):
    __tablename__ = 'companies'
    id = Column(Integer, primary_key=True)
    KharidarName = Column(String)
    KharidarLastNameSherkatName = Column(String)
    KharidarNationalCode = Column(String, unique=True)
    HCKharidarTypeCode = Column(String)


# Create the table in the database
Base.metadata.create_all(engine)


def create(**kwargs):
    Session = sessionmaker(bind=engine)
    company = Company(
        **kwargs
    )
    session = Session()
    session.add(company)
    session.commit()
    session.close()

def delete(company):
    Session = sessionmaker(bind=engine)
    session.delete(company)
    session.close()


def find(**kwargs):
    Session = sessionmaker(bind=engine)
    session = Session()
    result = session.query(Company).filter_by(**kwargs).first()
    session.close()
    return result


def fetchAll():
    Session = sessionmaker(bind=engine)
    session = Session()
    result = session.query(Company).all()
    session.close()
    return result
