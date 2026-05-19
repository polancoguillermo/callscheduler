import csv
import random
import os
from collections import defaultdict
from datetime import date, timedelta

# --- CONFIGURATION ---
HARD_START_DATE = "2026-07-01"

class Resident:
    def __init__(self, name, level, block_units, base_priority, weight):
        self.name = name
        self.level = level.lower()
        self.block_budgets = block_units 
        self.base_priority = base_priority 
        self.weight = float(weight)
        self.h1_units = 0
        self.h2_units = 0
        self.total_assigned_units = 0
        self.block_assignments = defaultdict(int)
        self.active_priorities = {} # {wk_idx: level}
        self._priority_dates = {} # {wk_idx: original_date_string} - for reporting
        self._assignments = []

    def remaining_in_block(self, block_num):
        return self.block_budgets.get(block_num, 0) - self.block_assignments[block_num]

class CallScheduler:
    def __init__(self, residents, n_blocks=13):
        self.residents = residents
        self.n_blocks = n_blocks
        self.n_weekends = n_blocks * 4
        base_date = date.fromisoformat(HARD_START_DATE)
        days_to_friday = (4 - base_date.weekday() + 7) % 7
        self.start_friday = base_date + timedelta(days=days_to_friday)
        self.schedule = [
            {d: {"senior": None, "junior": None} for d in ["friday", "saturday", "sunday"]} 
            for _ in range(self.n_weekends)
        ]
        self.violations = [] # Track failed priority accommodations

    def _can_assign(self, res, wk, day):
        block_num = (wk // 4) + 1
        u = {"friday": 1, "saturday": 2, "sunday": 1}[day]
        if res.remaining_in_block(block_num) < u: return False
        if self.schedule[wk][day][res.level]: return False
        wk_units = sum({"friday": 1, "saturday": 2, "sunday": 1}[d] for w, d in res._assignments if w == wk)
        if wk_units + u > 2: return False
        days_on = [d for w, d in res._assignments if w == wk]
        if day == "saturday" and ("friday" in days_on or "sunday" in days_on): return False
        if (day == "friday" or day == "sunday") and "saturday" in days_on: return False
        return True

    def _apply_fixed_assignment(self, res_name, target_date, role):
        delta = target_date - self.start_friday
        wk_idx = delta.days // 7
        day_map = {0: "friday", 1: "saturday", 2: "sunday"}
        day_str = day_map.get(delta.days % 7)
        if 0 <= wk_idx < self.n_weekends and day_str:
            res = next((r for r in self.residents if r.name == res_name), None)
            if res:
                u = {"friday": 1, "saturday": 2, "sunday": 1}[day_str]
                self.schedule[wk_idx][day_str][role] = res.name
                block_num = (wk_idx // 4) + 1
                if block_num <= 7: res.h1_units += u
                else: res.h2_units += u
                res.block_assignments[block_num] += u
                res.total_assigned_units += u
                res._assignments.append((wk_idx, day_str))
                return True
        return False

    def build(self):
        for wk in range(self.n_weekends):
            for role in ["senior", "junior"]:
                for day in ["friday", "saturday", "sunday"]:
                    if self.schedule[wk][day][role]: continue
                    pri_fri = (day == "sunday")
                    self._fill_slot(wk, day, role, prioritize_friday_person=pri_fri)
        
        # After building, check for violations
        self._check_priority_violations()

    def _fill_slot(self, wk, day, role, prioritize_friday_person=False):
        pool = [r for r in self.residents if r.level == role and self._can_assign(r, wk, day)]
        if not pool: return
        fri_person = self.schedule[wk]["friday"][role]
        
        def score_func(res):
            prio_penalty = res.active_priorities.get(wk, 0) * 1000000
            set_bonus = -500000 if prioritize_friday_person and fri_person == res.name else 0
            work_load_score = (res.total_assigned_units / res.weight) * 100
            return prio_penalty + set_bonus + work_load_score + random.uniform(0, 1)

        pool.sort(key=score_func)
        chosen = pool[0]
        u = {"friday": 1, "saturday": 2, "sunday": 1}[day]
        self.schedule[wk][day][role] = chosen.name
        block_num = (wk // 4) + 1
        if block_num <= 7: chosen.h1_units += u
        else: chosen.h2_units += u
        chosen.block_assignments[block_num] += u
        chosen.total_assigned_units += u
        chosen._assignments.append((wk, day))

    def _check_priority_violations(self):
        """Finds instances where a resident was assigned a shift on a priority date."""
        for res in self.residents:
            for wk_idx, original_date in res._priority_dates.items():
                # Check if resident was assigned any day in that weekend
                assigned_days = [d for w, d in res._assignments if w == wk_idx]
                if assigned_days:
                    self.violations.append({
                        "name": res.name,
                        "date": original_date,
                        "priority": res.active_priorities[wk_idx],
                        "assigned": ", ".join([d.capitalize() for d in assigned_days])
                    })

    def print_violation_report(self):
        print(f"\n{'!'*20} PRIORITY ACCOMMODATION REPORT {'!'*20}")
        if not self.violations:
            print("All priority requests were successfully accommodated!")
        else:
            print(f"{'Name':<15} | {'Requested Date':<15} | {'Level':<8} | {'Assigned Shift'}")
            print("-" * 65)
            for v in sorted(self.violations, key=lambda x: x['date']):
                print(f"{v['name']:<15} | {v['date']:<15} | P-{v['priority']:<6} | {v['assigned']}")
        print("-" * 65)

    def print_text_output(self, target_block=None):
        title = "FULL YEAR SCHEDULE" if target_block is None else f"BLOCK {target_block} SCHEDULE"
        print(f"\n{'='*20} {title} {'='*20}")
        print(f"{'Date':<12} | {'Day':<10} | {'Senior':<18} | {'Junior':<18}")
        print("-" * 65)
        for wk in range(self.n_weekends):
            if target_block and (wk // 4) + 1 != target_block: continue
            fri_date = self.start_friday + timedelta(weeks=wk)
            for i, day in enumerate(["friday", "saturday", "sunday"]):
                d = fri_date + timedelta(days=i)
                s = self.schedule[wk][day]["senior"] or "OPEN"
                j = self.schedule[wk][day]["junior"] or "OPEN"
                print(f"{d.strftime('%m/%d/%Y'):<12} | {day.capitalize():<10} | {s:<18} | {j:<18}")
            print("-" * 65)

def load_all_data():
    residents_map = {}
    try:
        with open('residents.csv', mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                budgets = {b: int(row.get(f'B{b}', 0)) for b in range(1, 14)}
                res = Resident(row['Name'], row['Level'], budgets, int(row.get('PriorityLevel', 0)), row.get('Weight', 1.0))
                residents_map[res.name] = res
    except: return None, None

    start_date = date.fromisoformat(HARD_START_DATE)
    first_friday = start_date + timedelta(days=(4 - start_date.weekday() + 7) % 7)

    try:
        with open('priorities.csv', mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                name = row['Name']
                if name in residents_map:
                    date_str = row['Date'].strip()
                    wk_idx = (date.fromisoformat(date_str) - first_friday).days // 7
                    if 0 <= wk_idx < 52: 
                        residents_map[name].active_priorities[wk_idx] = residents_map[name].base_priority
                        residents_map[name]._priority_dates[wk_idx] = date_str # Store for reporting
    except: pass
    return list(residents_map.values()), first_friday

if __name__ == "__main__":
    res_list, first_fri = load_all_data()
    if res_list:
        scheduler = CallScheduler(res_list)
        
        # Apply Fixed
        try:
            with open('fixed_schedule.csv', mode='r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    scheduler._apply_fixed_assignment(row['Name'], date.fromisoformat(row['Date']), row['Role'].lower())
        except: pass

        scheduler.build()
        
        # SHOW VIOLATIONS FIRST
        scheduler.print_violation_report()
        
        while True:
            print("\n--- VIEW OPTIONS ---")
            print("1. View Full Year")
            print("2. View Specific Block")
            print("3. Exit")
            choice = input("Select (1/2/3): ")
            if choice == "1": scheduler.print_text_output()
            elif choice == "2":
                try:
                    b_num = int(input("Enter Block # (1-13): "))
                    if 1 <= b_num <= 13: scheduler.print_text_output(target_block=b_num)
                except: pass
            elif choice == "3": break
