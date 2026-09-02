"""认证与权限装饰器"""
from functools import wraps
import json
from flask import jsonify, redirect, url_for, request, flash, abort
from flask_login import current_user
from sqlalchemy import or_
from models import Employee


def role_required(*roles):
    """角色权限装饰器: @role_required('admin','hr')"""
    def decorator(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            if not current_user.is_authenticated:
                # API 返回 401, 页面跳登录
                if request.path.startswith('/api/'):
                    return jsonify({'error': '未登录'}), 401
                return redirect(url_for('login'))
            if current_user.role not in roles:
                if request.path.startswith('/api/'):
                    return jsonify({'error': '权限不足'}), 403
                abort(403)
            return fn(*args, **kwargs)
        return wrapper
    return decorator


# 模块定义：key -> 显示名（与账号管理中的权限项一一对应）
MODULES = {
    'employee': '员工信息管理',
    'insurance': '五险一金管理',
    'salary': '薪酬绩效管理',
    'todos': '待办事项管理',
    'logs': '操作记录查询',
    'users': '账号管理',
    'dicts': '代码字段维护',
    'backup': '数据备份',
}


def _deny(msg='权限不足'):
    """统一拒绝：API 返回 403 JSON，页面 abort(403)"""
    if request.path.startswith('/api/'):
        return jsonify({'error': msg}), 403
    abort(403)


def parse_depts(s):
    """emp_depts 文本 -> 客户单位列表（支持 JSON 数组或逗号分隔）"""
    if not s:
        return []
    try:
        v = json.loads(s)
        if isinstance(v, list):
            return [str(x) for x in v if x]
    except Exception:
        pass
    return [x.strip() for x in str(s).replace('，', ',').split(',') if x.strip()]


def can_view_module(module):
    """当前用户能否查看某模块（打开页面 / 读取列表与数据）。"""
    if not current_user.is_authenticated:
        return False
    if current_user.role == 'admin':
        return True
    if current_user.role == 'hr':
        # 账号管理无需管理员授权；其余模块按授权
        if module == 'users':
            return True
        return bool(getattr(current_user, 'can_' + module + '_view', False))
    if current_user.role == 'employee':
        # 员工可查看「员工信息管理」「五险一金管理」（按本人范围）、「待办事项管理」（个人待办）与「账号管理」（仅本人）
        return module in ('employee', 'insurance', 'todos', 'users')
    return False


def can_manage_module(module):
    """当前用户能否管理（新增/修改/删除/导入）某模块。"""
    if not current_user.is_authenticated:
        return False
    if current_user.role == 'admin':
        return True
    if current_user.role == 'hr':
        # 账号管理无需管理员授权：可管理本人及员工账号（具体范围在接口内校验）
        if module == 'users':
            return True
        return bool(getattr(current_user, 'can_' + module + '_manage', False))
    if current_user.role == 'employee':
        # 员工可管理「待办事项」（个人）与「账号管理」（仅本人）
        return module in ('todos', 'users')
    return False  # 员工不可管理其余模块


def can_view_employee(emp_id):
    """判断当前用户能否查看某员工：admin/hr 看全部；员工按范围（仅自己 / 按客户单位）。"""
    if current_user.role in ('admin', 'hr'):
        return True
    if current_user.employee_id and current_user.employee_id == emp_id:
        return True
    if getattr(current_user, 'emp_view_mode', 'self') == 'dept':
        depts = parse_depts(getattr(current_user, 'emp_depts', ''))
        emp = Employee.query.get(emp_id)
        if emp and emp.department in depts:
            return True
    return False


def employee_scope_filter(query):
    """对员工查询施加「员工角色」的查看范围过滤（admin/hr 不受影响）。"""
    if current_user.role != 'employee':
        return query
    mode = getattr(current_user, 'emp_view_mode', 'self')
    depts = parse_depts(getattr(current_user, 'emp_depts', ''))
    eid = current_user.employee_id
    if mode == 'dept' and depts:
        # 按客户单位查看：本部门员工 + 自己（与 can_view_employee 口径一致）
        cond = Employee.department.in_(depts)
        if eid:
            cond = or_(Employee.id == eid, cond)
        return query.filter(cond)
    # self / 兜底
    if eid:
        return query.filter(Employee.id == eid)
    return query.filter(Employee.id == -1)
