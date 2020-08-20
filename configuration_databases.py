from sqlalchemy import create_engine, Column, String, Integer, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship

# 创建对象基类
Base = declarative_base()


# 定义软件版本模型类
class SoftwareVersion(Base):
    __tablename__ = 'softwareversion'
    id = Column(Integer, primary_key=True, autoincrement=True)
    software_version = Column(String(20), unique=True)

    def __repr__(self):
        return "<SofwareVersion(software_version='%s')>" % self.software_version


# 定义外部版本模型类
class CustomerVersion(Base):
    __tablename__ = 'customerversion'
    id = Column(Integer, primary_key=True, autoincrement=True)
    customer_version = Column(String(10), unique=True)


# 定义厂商代码模型类
class VendorCode(Base):
    __tablename__ = 'vendorcode'
    id = Column(Integer, primary_key=True, autoincrement=True)
    vendor_code = Column(String(5), unique=True)


# 定义日期模型类
class SoftwareDate(Base):
    __tablename__ = 'softwaredate'
    id = Column(Integer, primary_key=True)
    software_date = Column(String(10), unique=True)


# 定义WorkOrderNo模型类
class WorkOrderNo(Base):
    __tablename__ = 'work_order_no'
    id = Column(Integer, primary_key=True, autoincrement=True)
    work_order_no = Column(String(15), unique=True, nullable=False, index=True)
    chipids = relationship("ChipId", back_populates="workorderno")


# 定义ApprovalNo模型类
class ApprovalNo(Base):
    __tablename__ = 'approval_no'
    id = Column(Integer, primary_key=True, autoincrement=True)
    approval_no = Column(String(12), unique=True, nullable=False, index=True)
    chipids = relationship("ChipId", back_populates="approvalno")


# 定义ProductCategory模型类
class ProductCategory(Base):
    __tablename__ = 'product_category'
    id = Column(Integer, primary_key=True, autoincrement=True)
    product_category = Column(String(20), unique=True,
                              nullable=False, index=True)
    chipids = relationship("ChipId", back_populates="productcategory")


# 定义ChipId模型类
class ChipId(Base):
    __tablename__ = 'chip_id'
    id = Column(Integer, primary_key=True, autoincrement=True)
    chip_id = Column(String(48), unique=True, nullable=False)
    asset_no = Column(String(22), unique=True, nullable=False)
    work_order_no_id = Column(Integer, ForeignKey('work_order_no.id'))
    approval_no_id = Column(Integer, ForeignKey('approval_no.id'))
    product_category_id = Column(Integer, ForeignKey('product_category.id'))
    workorderno = relationship("WorkOrderNo", back_populates="chipids")
    approvalno = relationship("ApprovalNo", back_populates="chipids")
    productcategory = relationship("ProductCategory", back_populates="chipids")


# 引擎配置
engine = create_engine('sqlite:///configuration.db')
engine_chip_id = create_engine(
    'mysql+pymysql://root:Dream123$@heypython.cn:3306/microblog')


# 定义初始化数据库函数
def init_db():
    # Base.metadata.create_all(engine)
    Base.metadata.create_all(engine_chip_id)


# 顶固删除数据库函数
def drop_db():
    # Base.metadata.drop_all(engine)
    Base.metadata.drop_all(engine_chip_id)


# drop_db()
# init_db()
Session = sessionmaker(bind=engine)
session = Session()
Session_Chip_Id = sessionmaker(bind=engine_chip_id)
session_chip_id = Session_Chip_Id()
