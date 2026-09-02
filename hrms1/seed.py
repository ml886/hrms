"""初始化账号与示例数据"""
from datetime import date, timedelta
from app import app, db
from models import User, Employee, Contract, InsuranceDetail, Todo, DictItem, Setting


def seed_settings():
    """写入默认系统设置（已存在则跳过）"""
    defaults = {'sys_title': '人力资源管理系统', 'logo': ''}
    for k, v in defaults.items():
        if not Setting.query.filter_by(key=k).first():
            db.session.add(Setting(key=k, value=v))
    db.session.commit()


def seed_dicts():
    """初始化代码字段（字典）"""
    dicts = [
        # 客户单位
        ('客户单位', '研发部', 1), ('客户单位', '人力资源部', 2), ('客户单位', '市场部', 3),
        ('客户单位', '财务部', 4), ('客户单位', '销售部', 5), ('客户单位', '行政部', 6),
        # 职位
        ('职位', '总经理', 1), ('职位', '客户单位经理', 2), ('职位', '工程师', 3),
        ('职位', '架构师', 4), ('职位', 'HR专员', 5), ('职位', '会计', 6),
        ('职位', '销售代表', 7), ('职位', '行政助理', 8),
        # 学历
        ('学历', '初中及以下', 1), ('学历', '高中/中专', 2), ('学历', '大专', 3),
        ('学历', '本科', 4), ('学历', '硕士', 5), ('学历', '博士', 6),
        # 政治面貌
        ('政治面貌', '中共党员', 1), ('政治面貌', '共青团员', 2),
        ('政治面貌', '群众', 3), ('政治面貌', '民主党派', 4),
        # 民族
        ('民族', '汉族', 1), ('民族', '壮族', 2), ('民族', '回族', 3),
        ('民族', '满族', 4), ('民族', '维吾尔族', 5), ('民族', '苗族', 6),
        ('民族', '彝族', 7), ('民族', '土家族', 8), ('民族', '其他', 99),
        # 婚姻状况
        ('婚姻状况', '未婚', 1), ('婚姻状况', '已婚', 2),
        ('婚姻状况', '离异', 3), ('婚姻状况', '丧偶', 4),
        # 合同类型
        ('合同类型', '固定期限', 1), ('合同类型', '无固定期限', 2),
        ('合同类型', '以完成一定工作任务为期限', 3),
    ]
    added = 0
    for cat, label, sort in dicts:
        if not DictItem.query.filter_by(category=cat, label=label).first():
            db.session.add(DictItem(category=cat, label=label, sort=sort))
            added += 1
    if added:
        db.session.commit()
    return added


def seed():
    with app.app_context():
        db.create_all()
        seed_settings()
        seed_dicts()
        if User.query.count() > 0:
            print('已存在数据，跳过初始化')
            return

        # 员工
        today = date.today()
        emps = [
            Employee(name='张三', id_card='110101199001011234', phone='13800001111',
                     department='研发部', position='工程师', hire_date=date(2023, 3, 1),
                     status='在职', gender='男', birthday=date(1990, 1, 1),
                     address='北京市海淀区'),
            Employee(name='李四', id_card='110102199202022345', phone='13800002222',
                     department='人力资源部', position='HR专员', hire_date=date(2022, 6, 15),
                     status='在职', gender='女', birthday=date(1992, 2, 2),
                     address='北京市朝阳区'),
            Employee(name='王五', id_card='110103198803033456', phone='13800003333',
                     department='市场部', position='经理', hire_date=date(2021, 1, 10),
                     status='在职', gender='男', birthday=date(1988, 3, 3)),
            Employee(name='赵六', id_card='110104199504044567', phone='13800004444',
                     department='研发部', position='架构师', hire_date=date(2024, 5, 20),
                     status='在职', gender='男', birthday=date(1995, 4, 4)),
            Employee(name='孙七', id_card='110105199805055678', phone='13800005555',
                     department='财务部', position='会计', hire_date=date(2022, 9, 1),
                     status='离职', gender='女', birthday=date(1998, 5, 5),
                     leave_date=date(2024, 12, 31)),
        ]
        db.session.add_all(emps)
        db.session.flush()

        # 账号
        admin = User(username='admin', name='系统管理员', role='admin')
        admin.set_password('admin123')
        hr = User(username='hr', name='人事部', role='hr')
        hr.set_password('hr123')
        emp_user = User(username='emp', name='张三(员工)', role='employee',
                        employee_id=emps[0].id)
        emp_user.set_password('emp123')
        db.session.add_all([admin, hr, emp_user])

        # 合同（含即将到期和已到期）
        contracts = [
            Contract(employee_id=emps[0].id, contract_no='HT-2023-001',
                     start_date=date(2023, 3, 1),
                     end_date=today + timedelta(days=10),  # 即将到期
                     contract_type='固定期限', status='生效'),
            Contract(employee_id=emps[1].id, contract_no='HT-2022-002',
                     start_date=date(2022, 6, 15),
                     end_date=today + timedelta(days=120),
                     contract_type='固定期限', status='生效'),
            Contract(employee_id=emps[2].id, contract_no='HT-2021-003',
                     start_date=date(2021, 1, 10),
                     end_date=today - timedelta(days=5),  # 已过期
                     contract_type='固定期限', status='生效'),
            Contract(employee_id=emps[3].id, contract_no='HT-2024-004',
                     start_date=date(2024, 5, 20),
                     end_date=date(2027, 5, 19),
                     contract_type='固定期限', status='生效'),
            Contract(employee_id=emps[4].id, contract_no='HT-2022-005',
                     start_date=date(2022, 9, 1),
                     end_date=date(2025, 8, 31),
                     contract_type='固定期限', status='终止'),
        ]
        db.session.add_all(contracts)

        # 五险一金明细（用于汇总）
        details = []
        for emp in emps[:4]:  # 在职的4人
            details.append(InsuranceDetail(
                employee_id=emp.id, period='2026-07', base=10000,
                pension_emp=1600, pension_per=800,
                medical_emp=1000, medical_per=200,
                unemployment_emp=50, unemployment_per=50,
                injury_emp=30, maternity_emp=80,
                fund_emp=1200, fund_per=1200
            ))
            # 不同基数
            det2 = InsuranceDetail(
                employee_id=emp.id, period='2026-06', base=9500,
                pension_emp=1520, pension_per=760,
                medical_emp=950, medical_per=190,
                unemployment_emp=47.5, unemployment_per=47.5,
                injury_emp=28.5, maternity_emp=76,
                fund_emp=1140, fund_per=1140
            )
            details.append(det2)
        db.session.add_all(details)

        # 待办
        todos = [
            Todo(user_id=2, title='7月社保明细导入', due_date=today, priority='高'),
            Todo(user_id=2, title='员工合同续签面谈', due_date=today + timedelta(days=3),
                priority='高', remind_at=today + timedelta(days=1)),
            Todo(user_id=1, title='审批本月薪资调整', due_date=today - timedelta(days=1),
                priority='中'),
            Todo(user_id=3, title='更新个人联系方式', due_date=today + timedelta(days=7),
                priority='低'),
        ]
        db.session.add_all(todos)

        db.session.commit()
        print('初始化完成')
        print('账号: admin/admin123 | hr/hr123 | emp/emp123')


if __name__ == '__main__':
    seed()
