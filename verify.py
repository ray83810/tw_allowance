import os
import sys
import openpyxl
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

folder = r'c:\Users\User\Downloads\新增資料夾'

# Canonical names
CANONICAL_NAMES = [
    'Alex Chen', 'Amber Wang', 'Evan Liu', 'Howard Chen', 
    'Jacky Lee', 'Jian Kai Ding', 'Molly Song', 'Rex Liao', 'Sherry Lin'
]

def get_canonical_name(name):
    if not name:
        return ''
    name_clean = str(name).strip().replace(' ', '').lower()
    if 'ding' in name_clean and ('kai' in name_clean or 'jian' in name_clean or 'jlian' in name_clean):
        return 'Jian Kai Ding'
    
    # Try match
    for cname in CANONICAL_NAMES:
        c_clean = cname.replace(' ', '').lower()
        if name_clean == c_clean:
            return cname
        
        # Check parts
        c_parts = cname.lower().split()
        n_parts = str(name).lower().strip().split()
        if c_parts[-1] == n_parts[-1]:  # Same last name
            # Check overlap
            c_first = c_parts[0]
            n_first = n_parts[0]
            if c_first in n_first or n_first in c_first:
                return cname
    return str(name).strip()

def test_verification():
    # 1. Load data
    leave_wb = openpyxl.load_workbook(os.path.join(folder, '2026_請假申請表_Soluto_&_Care_May.xlsx'), data_only=True)
    ot_wb = openpyxl.load_workbook(os.path.join(folder, '2026_加班申請表_Soluto_&_Care_May.xlsx'), data_only=True)
    sched_wb = openpyxl.load_workbook(os.path.join(folder, '2026_排班表_May.xlsx'), data_only=True)
    target_wb = openpyxl.load_workbook(os.path.join(folder, '2026.5總請假統計.xlsx'), data_only=True)

    # Parse schedule employees
    sched_sheet = sched_wb['202605']
    sched_employees = []
    for r in range(3, sched_sheet.max_row+1):
        val = sched_sheet.cell(r, 1).value
        shift = sched_sheet.cell(r, 2).value
        if val and shift and any(c.isdigit() for c in str(shift)):
            sched_employees.append(get_canonical_name(val))
    sched_employees = sorted(list(set(sched_employees)))
    print('Schedule employees:', sched_employees)

    # Parse leave data
    leave_sheet = leave_wb['請假原始檔']
    leave_records = []
    for r in range(2, leave_sheet.max_row+1):
        applicant = leave_sheet.cell(r, 6).value
        ltype = leave_sheet.cell(r, 7).value
        mval = leave_sheet.cell(r, 8).value
        days = leave_sheet.cell(r, 11).value
        if applicant and ltype:
            name = get_canonical_name(applicant)
            if mval:
                if isinstance(mval, str):
                    m = int(mval.split('-')[1])
                else:
                    m = mval.month
            else:
                m = 0
            leave_records.append({'name': name, 'type': ltype, 'month': m, 'days': float(days or 0)})

    # Parse OT data
    ot_sheet = ot_wb['加班原始數據']
    ot_records = []
    for r in range(2, ot_sheet.max_row+1):
        applicant = ot_sheet.cell(r, 6).value
        date_val = ot_sheet.cell(r, 8).value
        hours = ot_sheet.cell(r, 12).value
        approved = ot_sheet.cell(r, 16).value
        if applicant and date_val:
            name = get_canonical_name(applicant)
            if approved == 'Approved':
                # Date parsing
                if isinstance(date_val, datetime):
                    dt_str = date_val.strftime('%Y-%m-%d')
                else:
                    dt_str = str(date_val).split()[0]
                ot_records.append({'name': name, 'date': dt_str, 'hours': float(hours or 0)})

    # Calculate leave stats for May
    may_leave_stats = {}
    for r in leave_records:
        if r['month'] == 5:
            may_leave_stats.setdefault(r['name'], {})
            may_leave_stats[r['name']][r['type']] = may_leave_stats[r['name']].get(r['type'], 0.0) + r['days']

    print('\nCalculated May Leave Stats:')
    for name in sorted(may_leave_stats.keys()):
        print(f'  {name}: {may_leave_stats[name]}')

    # Compare with 2026.5總請假統計.xlsx May sheet
    print('\nComparing with Target May sheet:')
    target_may_sheet = target_wb['May']
    for r in range(4, target_may_sheet.max_row+1):
        emp_name = target_may_sheet.cell(r, 3).value
        if emp_name:
            cname = get_canonical_name(emp_name)
            target_vals = {
                'Annual Leave特別休假': float(target_may_sheet.cell(r, 4).value or 0),
                'Asurion Leave亞勝假期': float(target_may_sheet.cell(r, 5).value or 0),
                'Sick Leave病假': float(target_may_sheet.cell(r, 6).value or 0),
                'Menstrual Leave病假（生理假)': float(target_may_sheet.cell(r, 7).value or 0),
                'Personal Leave事假': float(target_may_sheet.cell(r, 8).value or 0),
                'Official Leave公假': float(target_may_sheet.cell(r, 9).value or 0),
                'Marriage Leave婚假': float(target_may_sheet.cell(r, 10).value or 0),
                'Bereavement Leave喪假': float(target_may_sheet.cell(r, 11).value or 0),
                'Family Care Leave家庭照顧假': float(target_may_sheet.cell(r, 12).value or 0),
            }
            calc_vals = may_leave_stats.get(cname, {})
            diffs = []
            for k, tv in target_vals.items():
                cv = calc_vals.get(k, 0.0)
                if k == 'Official Leave公假':
                    cv = sum(v for kt, v in calc_vals.items() if 'Official Leave公假' in kt)
                if cv != tv:
                    diffs.append(f'{k}: calc={cv} vs target={tv}')
            if diffs:
                print(f'  [DIFF] {cname}: {", ".join(diffs)}')
            else:
                print(f'  [OK] {cname} matches target May sheet!')

    # Calculate OT stats
    calculated_ot = {}
    for r in ot_records:
        calculated_ot.setdefault(r['name'], {})
        calculated_ot[r['name']][r['date']] = calculated_ot[r['name']].get(r['date'], 0.0) + r['hours']

    print('\nCalculated OT Stats:')
    for name in sorted(calculated_ot.keys()):
        print(f'  {name}: {calculated_ot[name]}')

    # Compare OT with 2026_加班申請表_Soluto_&_Care_May.xlsx sheet 加班整理
    print('\nComparing OT with Target加班整理:')
    ot_target_sheet = ot_wb['加班整理']
    # Dates in row 4, cols 2 to max_column-1
    target_dates = []
    for col in range(2, ot_target_sheet.max_column):
        val = ot_target_sheet.cell(4, col).value
        if isinstance(val, datetime):
            target_dates.append((col, val.strftime('%Y-%m-%d')))
        elif val and val != '加總' and val != '總計':
            target_dates.append((col, str(val).split()[0]))
            
    # Check rows 5 onwards for employees
    for r in range(5, ot_target_sheet.max_row):
        emp_name = ot_target_sheet.cell(r, 1).value
        if emp_name and emp_name != '加總' and emp_name != '總計':
            cname = get_canonical_name(emp_name)
            diffs = []
            for col_idx, date_str in target_dates:
                target_val = float(ot_target_sheet.cell(r, col_idx).value or 0)
                calc_val = calculated_ot.get(cname, {}).get(date_str, 0.0)
                if target_val != calc_val:
                    diffs.append(f'{date_str}: calc={calc_val} vs target={target_val}')
            if diffs:
                print(f'  [DIFF] {cname}: {", ".join(diffs)}')
            else:
                print(f'  [OK] {cname} matches target OT sheet!')

    # Compare 統計表 cumulative numbers
    print('\nComparing 統計表 cumulative values:')
    target_stat_sheet_name = [s for s in target_wb.sheetnames if '統計表' in s][0]
    target_stat_sheet = target_wb[target_stat_sheet_name]
    
    # Cumulative stats from leave records (Jan-May)
    cumulative_leaves = {}
    for r in leave_records:
        if r['month'] <= 5:  # up to May
            cumulative_leaves.setdefault(r['name'], {})
            cumulative_leaves[r['name']][r['type']] = cumulative_leaves[r['name']].get(r['type'], 0.0) + r['days']

    for r in range(4, target_stat_sheet.max_row+1):
        emp_name = target_stat_sheet.cell(r, 1).value
        if emp_name:
            cname = get_canonical_name(emp_name)
            # Col B = PTO, Col C = PTO-AL
            # Col P = 已休, Col Q = 剩餘
            # Col R = 病假, Col S = 生理假, Col T = paid SL, Col U = paid SL-M
            # Col V = PL, Col W = LOA, Col X = ML, Col Y = BL, Col Z = FL
            target_pto_used = float(target_stat_sheet.cell(r, 16).value or 0)  # Col P
            target_pto_remain = float(target_stat_sheet.cell(r, 17).value or 0)  # Col Q
            target_sl = float(target_stat_sheet.cell(r, 20).value or 0)  # Col T (paid SL)
            target_slm = float(target_stat_sheet.cell(r, 21).value or 0)  # Col U (paid SL-M)
            target_pl = float(target_stat_sheet.cell(r, 22).value or 0)  # Col V (PL)
            target_loa = float(target_stat_sheet.cell(r, 23).value or 0)  # Col W (LOA)

            # Calc pto used
            emp_leaves = cumulative_leaves.get(cname, {})
            calc_pto_used = emp_leaves.get('Annual Leave特別休假', 0.0) + emp_leaves.get('Asurion Leave亞勝假期', 0.0)
            
            # Calc other leaves
            calc_sl_raw = emp_leaves.get('Sick Leave病假', 0.0)
            calc_slm_raw = emp_leaves.get('Menstrual Leave病假（生理假)', 0.0)
            calc_pl = emp_leaves.get('Personal Leave事假', 0.0)
            calc_loa = sum(v for k, v in emp_leaves.items() if 'Official Leave公假' in k)
            
            # Paid SL / SL-M logic
            calc_sl = calc_sl_raw + max(calc_slm_raw - 3.0, 0.0)
            calc_slm = min(calc_slm_raw, 3.0)

            diffs = []
            if calc_pto_used != target_pto_used:
                diffs.append(f'PTO Used: calc={calc_pto_used} vs target={target_pto_used}')
            if calc_sl != target_sl:
                diffs.append(f'SL: calc={calc_sl} vs target={target_sl}')
            if calc_slm != target_slm:
                diffs.append(f'SL-M: calc={calc_slm} vs target={target_slm}')
            if calc_pl != target_pl:
                diffs.append(f'PL: calc={calc_pl} vs target={target_pl}')
            if calc_loa != target_loa:
                diffs.append(f'LOA: calc={calc_loa} vs target={target_loa}')
                
            if diffs:
                print(f'  [DIFF] {cname}: {", ".join(diffs)}')
            else:
                print(f'  [OK] {cname} matches target cumulative stats!')

if __name__ == '__main__':
    test_verification()
