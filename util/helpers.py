class Shift:
    def __init__(self, time, hours, persons_required, shift_type, day, date):
        self.time = time
        self.hours = hours
        self.persons_required = persons_required
        self.shift_type = shift_type
        self.day = day
        self.date = date
        self.non_sunday_hours = 0 

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

    def __repr__(self):
        return f"Person(name={self.name}, availability={self.availability}, assigned_shifts={self.assigned_shifts})"

    def __str__(self):
        assigned_shifts_str = ", ".join([str(shift) for shift in self.assigned_shifts])
        return f"Person: {self.name}, Availability: {self.availability}, Assigned Shifts: {assigned_shifts_str}"


class Schedule:
    def __init__(self):
        self.shifts = []
        self.people = {}

    def get_total_availability(self):
        total_available = [0] * len(self.shifts)
        for shift_idx in range(len(self.shifts)):
            available_count = sum(
                person.availability[shift_idx] for person in self.people.values()
            )
            total_available[shift_idx] = available_count
        return total_available
    
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
