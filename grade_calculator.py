import math
def get_grade(score):
    if score < 0 or score > 100:
        raise ValueError("Score must be between 0 and 100.")

    if score >= 90:
        return "A"
    elif score >= 80:
        return "B"
    elif score >= 70:
        return "C"
    elif score >= 60:
        return "D"
    else:
        return "F"


while True:
    try:
        user_input = input("Enter your score (0-100): ").strip()

        # Handle empty input
        if not user_input:
            print("Error: Score cannot be empty.")
            continue

        # Convert input to number
        score = float(user_input)

        # Handle NaN and infinity
        if not math.isfinite(score):
            print("Error: Please enter a valid finite number.")
            continue

        # Validate range
        if score < 0 or score > 100:
            print("Error: Score must be between 0 and 100.")
            continue

        grade = get_grade(score)

        print(f"Mark: {score:g} -> Grade: {grade}")
        break

    except ValueError:
        print("Error: Please enter a valid number.")