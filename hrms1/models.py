"""HRMS 数据模型层 - SQLAlchemy ORM"""
import json
import contextvars
from datetime import date, datetime
import sqlalchemy as sa
from flask_sqlalchemy import SQLAlchemy
from flask_sqlalchemy.session import Session as _FSASession
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

# 多公司（独立库）动态绑定：按请求上下文选择当前公司引擎（contextvar 保证请求级隔离，无并发竞态）
current_company_engine = contextvars.ContextVar('hrms_company_engine', default=None)
# 未登录 / 未选公司时的兜底引擎（占位库 company_core.db，结构与公司库一致但为空）。
# 供 inject_globals 等「任何页面都会查询 Setting」的场景使用，避免 no such table 报错。
DEFAULT_COMPANY_ENGINE = None


class MultiTenantSession(_FSASession):
    """覆写 get_bind：当模型/表的 bind_key=='company' 时，返回当前请求上下文里的公司引擎，
    否则走默认引擎。从而实现「同一套代码、每公司独立库文件」。

    注意：本项目的公司表是在定义后才通过 __bind_key__ 挂到单一默认 metadata 上的，
    因此不能直接读 FSA 的 metadata.info['bind_key']（那里恒为 None），
    而要从「模型类属性 __bind_key__」或「表级 info['bind_key']」取标识。
    """
    def get_bind(self, mapper=None, clause=None, bind=None, **kwargs):
        if bind is not None:
            return bind
        bk = None
        if mapper is not None:
            try:
                bk = getattr(mapper.class_, '__bind_key__', None)
            except Exception:
                bk = None
        elif clause is not None:
            try:
                tables = clause.get_tables()
            except Exception:
                tables = []
            try:
                for t in tables:
                    if getattr(t, 'info', {}).get('bind_key') == 'company':
                        bk = 'company'
                        break
            except Exception:
                bk = None
        if bk == 'company':
            # 已选公司 -> 当前公司引擎；未选/未登录 -> 占位引擎兜底
            eng = current_company_engine.get() or DEFAULT_COMPANY_ENGINE
            if eng is not None:
                return eng
        return super().get_bind(mapper=mapper, clause=clause, bind=bind, **kwargs)


db = SQLAlchemy(session_options={'class_': MultiTenantSession})


class User(UserMixin, db.Model):
    """系统用户/账号"""
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    name = db.Column(db.String(64), nullable=False)
    role = db.Column(db.String(16), nullable=False, default='employee')  # admin / hr / employee
    # HR 模块权限：查看 / 管理（仅 hr 角色生效；admin 恒为全部，employee 忽略）
    can_employee_view = db.Column(db.Boolean, default=True)
    can_employee_manage = db.Column(db.Boolean, default=True)
    can_insurance_view = db.Column(db.Boolean, default=True)
    can_insurance_manage = db.Column(db.Boolean, default=True)
    can_todos_view = db.Column(db.Boolean, default=True)
    can_todos_manage = db.Column(db.Boolean, default=True)
    can_logs_view = db.Column(db.Boolean, default=True)
    can_logs_manage = db.Column(db.Boolean, default=True)
    can_salary_view = db.Column(db.Boolean, default=True)
    can_salary_manage = db.Column(db.Boolean, default=True)
    can_users_view = db.Column(db.Boolean, default=True)
    can_users_manage = db.Column(db.Boolean, default=True)
    can_dicts_view = db.Column(db.Boolean, default=True)
    can_dicts_manage = db.Column(db.Boolean, default=True)
    can_backup_view = db.Column(db.Boolean, default=True)
    can_backup_manage = db.Column(db.Boolean, default=True)
    # 员工查看范围：self=仅看自己，dept=按客户单位查看（emp_depts 多选）
    emp_view_mode = db.Column(db.String(8), default='self')
    emp_depts = db.Column(db.Text, default='')  # JSON 客户单位列表
    # 关联员工（逻辑关联，非外键：员工数据存于公司库，与全局用户库跨库）
    employee_id = db.Column(db.Integer, nullable=True)
    created_at = db.Column(db.Date, default=date.today)

    def set_password(self, pw):
        self.password_hash = generate_password_hash(pw)

    def check_password(self, pw):
        return check_password_hash(self.password_hash, pw)


class Employee(db.Model):
    """员工台账（入在离）— 公司库（每公司独立）"""
    __tablename__ = 'employees'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(64), nullable=False)
    id_card = db.Column(db.String(32), nullable=False)  # 身份证号
    phone = db.Column(db.String(32))
    department = db.Column(db.String(64))
    position = db.Column(db.String(64))
    hire_date = db.Column(db.Date, nullable=False)  # 入职日期
    status = db.Column(db.String(16), nullable=False, default='在职')  # 在职/离职
    leave_date = db.Column(db.Date, nullable=True)  # 离职日期
    gender = db.Column(db.String(8))
    birthday = db.Column(db.Date, nullable=True)
    address = db.Column(db.String(256))
    archive_no = db.Column(db.String(64))  # 档案编号
    # ---- 个人详细信息 ----
    native_place = db.Column(db.String(64))   # 籍贯
    ethnicity = db.Column(db.String(16))      # 民族
    education = db.Column(db.String(32))      # 学历
    school = db.Column(db.String(128))        # 毕业院校
    political_status = db.Column(db.String(16))  # 政治面貌
    marital_status = db.Column(db.String(16))    # 婚姻状况
    email = db.Column(db.String(128))         # 邮箱
    emergency_contact = db.Column(db.String(64))  # 紧急联系人
    emergency_phone = db.Column(db.String(32))    # 紧急联系电话
    remark = db.Column(db.String(512))           # 备注（离职/在职说明等）
    photo = db.Column(db.String(256))            # 个人证件照（存储文件名）
    # ---- 参保信息（参保时间 + 个人编号，按险种）----
    pension_enroll_date = db.Column(db.Date, nullable=True)   # 养老保险 参保时间
    pension_person_no = db.Column(db.String(32))              # 养老保险 个人编号
    injury_enroll_date = db.Column(db.Date, nullable=True)    # 工伤保险 参保时间
    injury_person_no = db.Column(db.String(32))               # 工伤保险 个人编号
    medical_enroll_date = db.Column(db.Date, nullable=True)   # 医疗保险 参保时间
    medical_person_no = db.Column(db.String(32))              # 医疗保险 个人编号
    maternity_enroll_date = db.Column(db.Date, nullable=True)  # 生育保险 参保时间
    maternity_person_no = db.Column(db.String(32))             # 生育保险 个人编号
    unemployment_enroll_date = db.Column(db.Date, nullable=True)  # 失业保险 参保时间
    unemployment_person_no = db.Column(db.String(32))         # 失业保险 个人编号
    fund_enroll_date = db.Column(db.Date, nullable=True)      # 住房公积金 参保时间
    fund_person_no = db.Column(db.String(32))                 # 住房公积金 个人编号
    created_at = db.Column(db.Date, default=date.today)

    contracts = db.relationship('Contract', backref='employee', lazy='dynamic')
    insurance_details = db.relationship('InsuranceDetail', backref='employee', lazy='dynamic')
    transfer_records = db.relationship('TransferRecord', backref='employee',
                                       order_by='TransferRecord.id.desc()',
                                       lazy='dynamic')


class Contract(db.Model):
    """劳动合同 — 公司库"""
    __tablename__ = 'contracts'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    contract_no = db.Column(db.String(64))  # 档案编号（合同沿用员工档案编号）
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=True)  # 无固定期限可空
    contract_type = db.Column(db.String(32), default='固定期限')  # 固定期限/无固定期限/以完成一定工作任务为期限
    status = db.Column(db.String(16), default='生效')  # 生效/到期/终止
    renew_count = db.Column(db.Integer, default=0)  # 续签次数
    remark = db.Column(db.String(256))
    created_at = db.Column(db.Date, default=date.today)

    renews = db.relationship('ContractRenew', backref='contract',
                             order_by='ContractRenew.renew_date.desc()',
                             lazy='dynamic')

    def days_to_expiry(self):
        """距到期天数；无固定期限返回 None"""
        if not self.end_date:
            return None
        return (self.end_date - date.today()).days


class ContractRenew(db.Model):
    """合同续签记录 — 公司库"""
    __tablename__ = 'contract_renews'
    id = db.Column(db.Integer, primary_key=True)
    contract_id = db.Column(db.Integer, db.ForeignKey('contracts.id'), nullable=False)
    renew_date = db.Column(db.Date, nullable=False)  # 续签日期
    old_end_date = db.Column(db.Date, nullable=True)  # 原到期日
    new_end_date = db.Column(db.Date, nullable=True)  # 新到期日
    remark = db.Column(db.String(256))  # 续签说明
    operator = db.Column(db.String(64))  # 操作人
    created_at = db.Column(db.Date, default=date.today)


class Attachment(db.Model):
    """员工附件管理 — 公司库"""
    __tablename__ = 'attachments'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    filename = db.Column(db.String(256), nullable=False)  # 存储文件名
    orig_name = db.Column(db.String(256), nullable=False)  # 原始文件名
    size = db.Column(db.Integer, default=0)  # 文件大小(字节)
    category = db.Column(db.String(32), default='其他')  # 分类
    uploader = db.Column(db.String(64))  # 上传人
    created_at = db.Column(db.Date, default=date.today)

    employee = db.relationship('Employee', backref='attachments')


class InsuranceDetail(db.Model):
    """五险一金明细（按员工按月）— 公司库"""
    __tablename__ = 'insurance_details'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    period = db.Column(db.String(7), nullable=False, index=True)  # YYYY-MM
    base = db.Column(db.Float, default=0)  # 缴费基数
    person_no = db.Column(db.String(32), default='')  # 人员编号
    remark = db.Column(db.String(512), default='')    # 备注

    pension_emp = db.Column(db.Float, default=0)   # 养老 单位
    pension_per = db.Column(db.Float, default=0)   # 养老 个人
    medical_emp = db.Column(db.Float, default=0)   # 医疗 单位
    medical_per = db.Column(db.Float, default=0)   # 医疗 个人
    extra_medical_emp = db.Column(db.Float, default=0)  # 大额医疗补助 单位
    extra_medical_per = db.Column(db.Float, default=0)  # 大额医疗补助 个人
    unemployment_emp = db.Column(db.Float, default=0)  # 失业 单位
    unemployment_per = db.Column(db.Float, default=0)  # 失业 个人
    injury_emp = db.Column(db.Float, default=0)   # 工伤 单位
    maternity_emp = db.Column(db.Float, default=0)  # 生育 单位
    fund_base = db.Column(db.Float, default=0)     # 公积金缴费基数
    fund_rate = db.Column(db.Float, default=0)     # 缴费比例(%)
    fund_emp = db.Column(db.Float, default=0)      # 公积金 单位
    fund_per = db.Column(db.Float, default=0)      # 公积金 个人

    created_at = db.Column(db.Date, default=date.today)

    __table_args__ = (db.UniqueConstraint('employee_id', 'period', name='uq_emp_period'),)


class SalaryRecord(db.Model):
    """工资汇总明细（按员工按月；数据由模版导入，列可自定义显示）— 公司库"""
    __tablename__ = 'salary_records'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=True)
    period = db.Column(db.String(7), nullable=False, index=True)  # YYYY-MM
    name = db.Column(db.String(32), default='')        # 员工姓名（冗余，便于导出/离职后保留）
    id_card = db.Column(db.String(32), default='')     # 身份证号（匹配键）
    base_salary = db.Column(db.Float, default=0)      # 基本工资
    post_salary = db.Column(db.Float, default=0)      # 岗位工资
    performance_salary = db.Column(db.Float, default=0)  # 绩效工资
    allowance = db.Column(db.Float, default=0)        # 津贴补贴
    overtime_pay = db.Column(db.Float, default=0)     # 加班工资
    bonus = db.Column(db.Float, default=0)            # 奖金
    should_pay = db.Column(db.Float, default=0)       # 应发工资
    social_personal = db.Column(db.Float, default=0) # 社保个人
    fund_personal = db.Column(db.Float, default=0)   # 公积金个人
    tax = db.Column(db.Float, default=0)              # 个税
    deduct_total = db.Column(db.Float, default=0)    # 扣款合计
    net_pay = db.Column(db.Float, default=0)          # 实发工资
    client_unit = db.Column(db.String(128), default='')  # 客户单位（来自导入表/工资表表头单位名称）
    remark = db.Column(db.String(512), default='')    # 备注
    # 原始表格内容（按导入/校验表格实际列名与逐行值存储，用于查询页按原表内容展示）
    headers_json = db.Column(db.Text, default='')     # JSON 数组：导入/校验表格的实际列名（按原表顺序）
    values_json = db.Column(db.Text, default='')      # JSON 对象：本行各列实际值（列名 -> 值）
    created_at = db.Column(db.Date, default=date.today)

    __table_args__ = (db.UniqueConstraint('id_card', 'period', name='uq_salary_idcard_period'),)


class Todo(db.Model):
    """自定义待办事项"""
    __tablename__ = 'todos'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    company_id = db.Column(db.Integer, db.ForeignKey('companies.id'), nullable=True)  # 所属公司（按公司分开管理/提醒）
    title = db.Column(db.String(256), nullable=False)
    due_date = db.Column(db.Date, nullable=False)
    priority = db.Column(db.String(8), default='中')  # 高/中/低
    done = db.Column(db.Boolean, default=False)
    remind_at = db.Column(db.Date, nullable=True)  # 提醒日期（旧字段，保留兼容）
    remind_time = db.Column(db.String(8), nullable=True)  # 提醒时间 HH:MM（旧字段，保留兼容）
    remind_start = db.Column(db.DateTime, nullable=True)  # 提醒开始时间
    remind_end = db.Column(db.DateTime, nullable=True)    # 提醒结束时间
    remind_count = db.Column(db.Integer, default=1)      # 提醒次数
    reminded_count = db.Column(db.Integer, default=0)    # 已提醒次数
    created_at = db.Column(db.Date, default=date.today)

    user = db.relationship('User', backref='todos', foreign_keys=[user_id])


class OperationLog(db.Model):
    """系统操作日志/记录"""
    __tablename__ = 'operation_logs'
    id = db.Column(db.Integer, primary_key=True)
    operator = db.Column(db.String(64), nullable=False)      # 操作人账号
    operator_name = db.Column(db.String(64))                 # 操作人姓名
    module = db.Column(db.String(32), nullable=False)        # 模块（员工/合同/附件/待办/账号/字典/五险一金等）
    action = db.Column(db.String(32), nullable=False)        # 操作类型（新增/编辑/删除/转岗/离职/续签/上传/导入等）
    target = db.Column(db.String(128))                       # 操作对象描述
    detail = db.Column(db.String(512))                       # 详情说明
    ip = db.Column(db.String(64))                            # 操作IP
    created_at = db.Column(db.DateTime, default=lambda: datetime.now())  # 本地时间（非 UTC）


class TransferRecord(db.Model):
    """转岗/调动记录（可追溯历史）— 公司库"""
    __tablename__ = 'transfer_records'
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employees.id'), nullable=False)
    old_department = db.Column(db.String(64))
    new_department = db.Column(db.String(64))
    old_position = db.Column(db.String(64))
    new_position = db.Column(db.String(64))
    remark = db.Column(db.String(512))   # 转岗备注
    transfer_date = db.Column(db.Date, default=date.today)  # 转岗日期
    operator = db.Column(db.String(64))  # 操作人
    created_at = db.Column(db.Date, default=date.today)


class DictItem(db.Model):
    """代码字段（字典）维护：客户单位/职位/学历/政治面貌/民族/婚姻状况/合同类型等"""
    __tablename__ = 'dict_items'
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(32), nullable=False, index=True)  # 分类
    label = db.Column(db.String(64), nullable=False)  # 选项值
    sort = db.Column(db.Integer, default=0)  # 排序
    enabled = db.Column(db.Boolean, default=True)
    remark = db.Column(db.String(256))
    created_at = db.Column(db.Date, default=date.today)

    __table_args__ = (db.UniqueConstraint('category', 'label', name='uq_cat_label'),)


class Setting(db.Model):
    """系统设置键值对（系统名称、logo 文件名等）— 公司库（每公司独立）"""
    __tablename__ = 'settings'
    id = db.Column(db.Integer, primary_key=True)
    key = db.Column(db.String(64), unique=True, nullable=False, index=True)
    value = db.Column(db.Text, default='')

    @staticmethod
    def get(key, default=''):
        row = Setting.query.filter_by(key=key).first()
        return row.value if row else default

    @staticmethod
    def set(key, value):
        row = Setting.query.filter_by(key=key).first()
        if not row:
            row = Setting(key=key, value=value)
            db.session.add(row)
        else:
            row.value = value
        db.session.flush()


class PayslipTemplate(db.Model):
    """工资表生成固定模板库：命名保存，可多次使用；可带样式文件（字体/列宽/合并等）— 公司库"""
    __tablename__ = 'payslip_templates'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    columns_json = db.Column(db.Text, nullable=False, default='[]')  # 生成列定义 list
    file_blob = db.Column(db.LargeBinary, nullable=True)            # 带样式模板 xlsx 字节
    file_name = db.Column(db.String(255), nullable=True)
    headfoot_json = db.Column(db.Text, nullable=True)              # 表头/表脚模板（结构化 {header, footer}）
    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    def to_dict(self):
        hf = None
        if self.headfoot_json:
            try:
                hf = json.loads(self.headfoot_json)
            except Exception:
                hf = None
        # v2 表头表脚：{per_sheet:{name:{header_after,footer_after}}, tax:{...}}
        # 兼容旧格式 {header, footer}
        has_hf = False
        tax = None
        if isinstance(hf, dict):
            tax = hf.get('tax')
            if hf.get('per_sheet'):
                for v in hf['per_sheet'].values():
                    if v and (v.get('header_after') or v.get('footer_after')):
                        has_hf = True
                        break
            if hf.get('header_after') or hf.get('footer_after'):
                has_hf = True
            if not has_hf and (hf.get('header') or hf.get('footer')):
                has_hf = True
        return {
            'id': self.id, 'name': self.name,
            'columns': json.loads(self.columns_json) if self.columns_json else [],
            'has_file': bool(self.file_blob),
            'file_name': self.file_name,
            'headfoot': hf,
            'has_headfoot': has_hf,
            'tax': tax,
            'created_at': self.created_at.strftime('%Y-%m-%d %H:%M') if self.created_at else '',
            'updated_at': self.updated_at.strftime('%Y-%m-%d %H:%M') if self.updated_at else '',
        }


# 公司模型统一标记为公司隔离（每个公司一个 SQLite 文件，由 MultiTenantSession.get_bind 动态路由）。
# 因为表都挂在单一默认 metadata 上，FSA 的 metadata.info['bind_key'] 不可用，
# 故同时设置「类属性 __bind_key__」（供 mapper 路径）与「表级 info['bind_key']」（供 clause 路径）。
for _m in (Employee, Contract, ContractRenew, Attachment, InsuranceDetail,
           SalaryRecord, TransferRecord, PayslipTemplate, Setting, OperationLog, DictItem):
    _m.__bind_key__ = 'company'
    try:
        _m.__table__.info['bind_key'] = 'company'
    except Exception:
        pass


class Company(db.Model):
    """公司（账套）：每公司一套独立数据库，框架一致、数据独立。"""
    __tablename__ = 'companies'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(128), nullable=False)        # 公司名称
    code = db.Column(db.String(32), unique=True, nullable=False, index=True)  # 公司编码（用于库文件名）
    enabled = db.Column(db.Boolean, default=True)
    remark = db.Column(db.String(256))
    logo = db.Column(db.String(255), default='')   # 公司 Logo 文件名（存于 uploads/company_logos/）；空=用默认图
    created_at = db.Column(db.Date, default=date.today)

    def db_path(self):
        """该公司库文件路径（位于 BASE_DIR/companies/<code>.db）"""
        from flask import current_app
        base = current_app.config.get('HRMS_BASE_DIR', '')
        import os
        d = os.path.join(base, 'companies')
        os.makedirs(d, exist_ok=True)
        return os.path.join(d, f'{self.code}.db')


class UserCompany(db.Model):
    """用户—公司 权限关联（仅 hr/employee 生效；admin 恒拥有全部公司）。"""
    __tablename__ = 'user_companies'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False, index=True)
    company_id = db.Column(db.Integer, db.ForeignKey('companies.id'), nullable=False, index=True)
    created_at = db.Column(db.Date, default=date.today)

    __table_args__ = (db.UniqueConstraint('user_id', 'company_id', name='uq_user_company'),)


class SysSetting(db.Model):
    """全局系统设置（系统名称、登录页 Logo 等），存于主库 data.db，与具体公司无关。
    用于登录页品牌展示；左上角的公司切换区不再读取这些设置（已解耦为公司维度）。"""
    __tablename__ = 'sys_settings'
    id = db.Column(db.Integer, primary_key=True)
    key = db.Column(db.String(64), unique=True, nullable=False, index=True)
    value = db.Column(db.Text, default='')

    @staticmethod
    def get(key, default=''):
        row = SysSetting.query.filter_by(key=key).first()
        return row.value if row else default

    @staticmethod
    def set(key, value):
        row = SysSetting.query.filter_by(key=key).first()
        if not row:
            row = SysSetting(key=key, value=value)
            db.session.add(row)
        else:
            row.value = value
        db.session.flush()

