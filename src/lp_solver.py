import gurobipy as gp
import numpy as np
import pandas as pd

class LPSolver:
    def __init__(self, schedule, max_hours, sunday_quota):
        self.schedule = schedule
        self.max_hours = max_hours
        self.sunday_quota = sunday_quota
        self.model = gp.Model()
        self.model.setParam('TimeLimit', 100)
        self.A = {}  # Decision variables
        self.slack = {}  # Slack variables for unfilled shifts

    def setup_variables(self):
        for shift_idx, shift in enumerate(self.schedule.shifts):
            for person_name, person in self.schedule.people.items():
                self.A[(person_name, shift_idx)] = self.model.addVar(vtype=gp.GRB.BINARY, name=f"A_{person_name}_{shift_idx}")

            # Add slack variable to represent unfilled shifts (if available people are fewer than required)
            self.slack[shift_idx] = self.model.addVar(vtype=gp.GRB.INTEGER, lb=0, name=f"Slack_{shift_idx}")

    def set_objective(self):
        # Main objective composition
        objective = 0
        
        # Add core hour distribution objectives
        objective +=  self._build_hour_distribution_terms()
        
        # Add spread penalty for shift clustering
        objective +=  10 * self._build_shift_spread_penalty()
        
        # Add slack penalty for unfilled shifts
        objective += self._build_slack_penalty()
        
        self.model.setObjective(objective, gp.GRB.MINIMIZE)

    def _build_hour_distribution_terms(self):
        # Calculate hour distribution error terms
        total_available, total_regular, total_bonus = self._precompute_totals()
        error_terms = 0

        for person_name, person in self.schedule.people.items():
            assigned_regular, assigned_bonus = self._get_assigned_hours(person_name)
            err_reg, err_bonus = self._create_error_variables(
                person, assigned_regular, assigned_bonus, 
                total_available, total_regular, total_bonus
            )
            error_terms += err_reg**2 + 0.3*err_bonus**2

        return error_terms

    def _build_shift_spread_penalty(self):
        # Calculate shift spread penalty with decaying weights
        penalty = 0
        days_order = ["Maandag", "Dinsdag", "Woensdag", "Donderdag",
                    "Vrijdag", "Zaterdag", "Zondag"]
        
        for person_name in self.schedule.people:
            penalty += self._calculate_person_spread_penalty(person_name, days_order)
        
        return penalty

    def _calculate_person_spread_penalty(self, person_name, days_order):
        # Penalty calculation for a single person's shifts
        penalty = 0
        params = {
            'max_gap': 3,
            'coeff': 0.1,
            'weights': {0: 0, 1: 3, 2: 2, 3: 1}
        }

        for i, shift_i in enumerate(self.schedule.shifts):
            for j, shift_j in enumerate(self.schedule.shifts[i+1:], start=i+1):
                day_diff = self._calculate_day_gap(
                    shift_i.day, shift_j.day, days_order
                )
                
                if day_diff > params['max_gap']:
                    continue
                    
                penalty += params['coeff'] * params['weights'][day_diff] * \
                        self.A[(person_name, i)] * self.A[(person_name, j)]

        return penalty

    def _build_slack_penalty(self):
        # Penalty for unfilled shifts
        return 100000 * gp.quicksum(self.slack.values())

    def _precompute_totals(self):
        # Helper for total hour calculations
        return (
            sum(p.available_regular_hours for p in self.schedule.people.values()),
            sum(s.hours * s.persons_required for s in self.schedule.shifts),
            sum(s.bonus_hours * s.persons_required for s in self.schedule.shifts)
        )

    def _get_assigned_hours(self, person_name):
        # Calculate hours assigned to a person
        regular = gp.quicksum(
            self.A[(person_name, i)] * s.hours 
            for i, s in enumerate(self.schedule.shifts)
        )
        bonus = gp.quicksum(
            self.A[(person_name, i)] * s.bonus_hours 
            for i, s in enumerate(self.schedule.shifts)
        )
        return regular, bonus

    def _create_error_variables(self, person, assigned_reg, assigned_bonus, 
                            total_available, total_reg, total_bonus):
        # Create error terms for hour distribution
        if total_available > 0:
            exp_reg = (person.available_regular_hours/total_available) * total_reg
            exp_bonus = (person.available_regular_hours/total_available) * total_bonus
        else:
            exp_reg = exp_bonus = 0

        err_reg = self.model.addVar(lb=-gp.GRB.INFINITY)
        err_bonus = self.model.addVar(lb=-gp.GRB.INFINITY)
        
        self.model.addConstr(err_reg == exp_reg - assigned_reg)
        self.model.addConstr(err_bonus == exp_bonus - assigned_bonus)
        
        return err_reg, err_bonus

    def _calculate_day_gap(self, day1, day2, days_order):
        # Normalize input days
        day1 = str(day1).strip().title()
        day2 = str(day2).strip().title()
        days_order = [d.strip().title() for d in days_order]

        # Check for true NaN values (if using pandas)
        if pd.isna(day1) or pd.isna(day2):
            raise ValueError(f"NaN day detected: {day1} or {day2}")

        if day1 not in days_order or day2 not in days_order:
            raise ValueError(
                f"Invalid day(s): '{day1}' or '{day2}'. "
                f"Must be one of {days_order}"
            )

        idx1 = days_order.index(day1)
        idx2 = days_order.index(day2)
        return min(abs(idx2 - idx1), 7 - abs(idx2 - idx1))

    def apply_constraints(self):
        self._apply_shift_assignment_constraints()
        self._apply_availability_constraints()
        self._apply_no_night_to_morning_constraints()
        self._apply_no_evening_to_morning_constraints()
        self._apply_no_sunday_if_quota_not_met_constraints()
        self._apply_max_one_sunday_shift_constraints()
        self._apply_max_hours_constraints()

    def _apply_shift_assignment_constraints(self):
        for shift_idx, shift in enumerate(self.schedule.shifts):
            # Total assignments for each shift must equal the number of persons required for that shift, plus slack for unfilled shifts
            self.model.addConstr(
                sum(self.A[(person_name, shift_idx)] for person_name in self.schedule.people) + self.slack[shift_idx],
                gp.GRB.EQUAL, 
                shift.persons_required, 
                name=f"C2_ShiftAssignment_{shift_idx}"
            )

    def _apply_availability_constraints(self):
        for shift_idx, shift in enumerate(self.schedule.shifts):
            for person_name, person in self.schedule.people.items():
                self.model.addConstr(self.A[(person_name, shift_idx)] <= person.availability[shift_idx],
                                     name=f"C3_Availability_{person_name}_{shift_idx}")

    def _apply_no_night_to_morning_constraints(self):
        for person_name, person in self.schedule.people.items():
            prev_shift = None
            prev_type = None
            for shift_idx, shift in enumerate(self.schedule.shifts):
                if prev_shift is not None and prev_type == "Avond" and shift.shift_type == "Ochtend":
                    self.model.addConstr(self.A[(person_name, shift_idx)] + self.A[(person_name, prev_shift)] <= 1,
                                         name=f"C4_NoNightToMorning_{person_name}_{shift_idx}")
                prev_shift, prev_type = shift_idx, shift.shift_type

    def _apply_no_evening_to_morning_constraints(self):
        for person_name, person in self.schedule.people.items():
            prev_shift, prev_prev_shift = None, None
            prev_type, prev_prev_type = None, None
            for shift_idx, shift in enumerate(self.schedule.shifts):
                if shift.shift_type == "Avond":
                    if prev_type in ["Middag", "Ochtend"]:
                        self.model.addConstr(self.A[(person_name, prev_shift)] + self.A[(person_name, shift_idx)] <= 1,
                                             name=f"C5.1_NoAfternoonToEvening_{person_name}_{shift_idx}")
                    if prev_prev_type == "Ochtend":
                        self.model.addConstr(self.A[(person_name, prev_prev_shift)] + self.A[(person_name, shift_idx)] <= 1,
                                             name=f"C5.2_NoMorningToEvening_{person_name}_{shift_idx}")
                prev_prev_shift, prev_prev_type = prev_shift, prev_type
                prev_shift, prev_type = shift_idx, shift.shift_type

    def _apply_no_sunday_if_quota_not_met_constraints(self):
        for person_name, person in self.schedule.people.items():
            if person.non_sunday_hours < self.sunday_quota:
                for shift_idx, shift in enumerate(self.schedule.shifts):
                    if shift.day == "Zondag":
                        self.model.addConstr(self.A[(person_name, shift_idx)] == 0,
                                             name=f"C6_NoSundayIfQuotaNotMet_{person_name}_{shift_idx}")

    def _apply_max_one_sunday_shift_constraints(self):
        for person_name in self.schedule.people:
            self.model.addConstr(sum(self.A[(person_name, shift_idx)] for shift_idx, shift in enumerate(self.schedule.shifts) if shift.day == "Zondag") <= 1,
                                 name=f"C7_OneSundayShift_{person_name}")

    def _apply_max_hours_constraints(self):
        for person_name in self.schedule.people:
            self.model.addConstr(sum(self.A[(person_name, shift_idx)] * shift.hours for shift_idx, shift in enumerate(self.schedule.shifts)) <= self.max_hours,
                                 name=f"C8_MaxHours_{person_name}")

    def solve(self):
        self.model.optimize()
