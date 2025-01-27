class Shift:
    def __init__(self, time, hours, persons_required, shift_type, day, date):
        self.time = time
        self.hours = hours
        self.persons_required = persons_required
        self.shift_type = shift_type
        self.day = day
        self.date = date
        self.non_sunday_hours = 0
        self.bonus_hours = self._calculate_bonus()  # New attribute

    def _calculate_bonus(self):
        if self.day == "Zaterdag":
            return 0.4 * self.hours
        elif self.day == "Zondag":
            return 0.75 * self.hours
        elif self.shift_type == "Avond":
            return 0.4 * 4
        else:
            return 0.0

    def __repr__(self):
        return f"Shift(time={self.time}, hours={self.hours}, persons={self.persons_required}, shift_type={self.shift_type}, day={self.day}, date={self.date})"

    def __str__(self):
        return f"Shift on {self.date} ({self.day}) from {self.time}, {self.hours} hours, Type: {self.shift_type}, Needed: {self.persons_required}"


class Person:
    def __init__(self, name):
        self.name = name
        self.availability = []
        self.assigned_shifts = []
        self.expected_share = 0
        self.available_regular_hours = 0  # Total regular hours available
        self.available_bonus_hours = 0    # Total bonus hours available
        self.total_available_hours = 0 

    def __repr__(self):
        return f"Person(name={self.name}, availability={self.availability}, assigned_shifts={self.assigned_shifts})"

    def __str__(self):
        assigned_shifts_str = ", ".join([str(shift) for shift in self.assigned_shifts])
        return f"Person: {self.name}, Availability: {self.availability}, Assigned Shifts: {assigned_shifts_str}, Bonus Hours: {self.bonus_hours}"

class Schedule:
    def __init__(self):
        self.shifts = []
        self.people = {}

    def calculate_availability(self):
        """Precompute available regular, bonus, and total hours for each person."""
        for person in self.people.values():
            person.available_regular_hours = 0
            person.available_bonus_hours = 0
            person.total_available_hours = 0
            for shift_idx, shift in enumerate(self.shifts):
                if person.availability[shift_idx]:
                    person.available_regular_hours += shift.hours
                    person.available_bonus_hours += shift.bonus_hours
                    person.total_available_hours += shift.hours + shift.bonus_hours

    def get_total_available_regular(self):
        """Returns the sum of available regular hours across all people."""
        return sum(person.available_regular_hours for person in self.people.values())
    
    def calculate_non_sunday_hours(self):
        """Calculates and stores non-Sunday available hours for each person."""
        for person_name, person in self.people.items():
            person.non_sunday_hours = sum(
                shift.hours * person.availability[shift_idx]  
                for shift_idx, shift in enumerate(self.shifts)
                if shift.day != "Zondag"  # Exclude Sundays
            )

    def __repr__(self):
        return f"Schedule(shifts={self.shifts}, people={self.people})"

    def __str__(self):
        shifts_str = ", ".join(map(str, self.shifts))
        people_str = ", ".join(map(str, self.people.values()))
        return f"Schedule:\nShifts:\n{shifts_str}\n\nPeople:\n{people_str}"
