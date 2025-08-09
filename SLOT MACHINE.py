import time
import random
import string
import sys
from collections import Counter

choices = ("🍒", "🍋", "🍉", "🔔", "💎")
balance = ""
total_winnings = ""
v_score = 0
h_score = 0


def generate_symbols():
    random_symbols_one = random.choice(choices)
    random_symbols_two = random.choice(choices)
    random_symbols_three = random.choice(choices)
    random_symbols_four = random.choice(choices)
    random_symbols_five = random.choice(choices)
    total_symbols = random_symbols_one + random_symbols_two + random_symbols_three + random_symbols_four + random_symbols_five
    return total_symbols


def vertical_score(v_score):
    vertical_3_pairs = ""
    vertical_2_pairs_one_and_two = ""
    vertical_2_pairs_two_and_three = ""

    # how to compute values vertically
    for y in range(5):
        if slot_line_one[y] == slot_line_two[y] == slot_line_three[y]:  # line 1 , line 2 , line 3 - 3 PAIRS
            v_score += 3
            vertical_3_pairs += f" {slot_line_one[y]} : {slot_line_two[y]} : {slot_line_three[y]}"

        elif slot_line_one[y] == slot_line_two[y]:  # line 1 and line 2 - 2 PAIRS
            v_score += 2
            vertical_2_pairs_one_and_two += f" {slot_line_one[y]} : {slot_line_two[y]}"

        elif slot_line_two[y] == slot_line_three[y]:  # line 2 and line 3 - 2 PAIRS
            v_score += 2
            vertical_2_pairs_two_and_three += f" {slot_line_two[y]} : {slot_line_three[y]}"

    return v_score, vertical_3_pairs, vertical_2_pairs_one_and_two, vertical_2_pairs_two_and_three


def pattern_count_horizontal(run_length):
    if run_length == 2:
        return 2
    elif run_length == 3:
        return 3
    elif run_length == 4:
        return 5
    elif run_length == 5:
        return 10
    else:
        return 0


def horizontal_score_line_one(slot_line_one, h_score):  # horizontal score for line 1
    run_length = 1
    patterns_h1 = []
    h_score_line_one = ""
    for i in range(len(slot_line_one) - 1):
        if slot_line_one[i] == slot_line_one[i + 1]:
            run_length += 1
        else:
            if run_length >= 2:
                patterns_h1.append(f"{slot_line_one[i]} x{run_length}")
            h_score += pattern_count_horizontal(run_length)
            run_length = 1

    # append last run if >=2
    if run_length >= 2:
        patterns_h1.append(f"{slot_line_one[-1]} x{run_length}")
    h_score += pattern_count_horizontal(run_length)

    if patterns_h1:
        for pattern in patterns_h1:
            h_score_line_one += f" {pattern}"

    return h_score , h_score_line_one


def horizontal_score_line_two(slot_line_two, h_score):  # horizontal score for line 2
    run_length = 1
    patterns_h2 = []
    h_score_line_two = ""

    for i in range(len(slot_line_two) - 1):
        if slot_line_two[i] == slot_line_two[i + 1]:
            run_length += 1
        else:
            if run_length >= 2:
                patterns_h2.append(f"{slot_line_two[i]} x{run_length}")
            h_score += pattern_count_horizontal(run_length)
            run_length = 1

    # append last run if >=2
    if run_length >= 2:
        patterns_h2.append(f"{slot_line_two[-1]} x{run_length}")
    h_score += pattern_count_horizontal(run_length)

    if patterns_h2:
        for pattern in patterns_h2:
            h_score_line_two += f" {pattern}"

    return h_score, h_score_line_two


def horizontal_score_line_three(slot_line_three, h_score):  # horizontal score for line 3
    run_length = 1
    patterns_h3 = []
    h_score_line_three = ""
    for i in range(len(slot_line_three) - 1):
        if slot_line_three[i] == slot_line_three[i + 1]:
            run_length += 1
        else:
            if run_length >= 2:
                patterns_h3.append(f"{slot_line_three[i]} x{run_length}")
            h_score += pattern_count_horizontal(run_length)
            run_length = 1

    # append last run if >=2
    if run_length >= 2:
        patterns_h3.append(f"{slot_line_three[-1]} x{run_length}")
    h_score += pattern_count_horizontal(run_length)

    if patterns_h3:
        for pattern in patterns_h3:
            h_score_line_three += f" {pattern}"

    return h_score, h_score_line_three


# multiple assignment
total_symbols = generate_symbols()
slot_line_one = generate_symbols()
slot_line_two = generate_symbols()
slot_line_three = generate_symbols()

# horizontal score
h_score, h_score_line_one = horizontal_score_line_one(slot_line_one, h_score)
h_score, h_score_line_two = horizontal_score_line_two(slot_line_two, h_score)
h_score, h_score_line_three = horizontal_score_line_three(slot_line_three, h_score)

# vertical score
v_score, vertical_3_pairs, vertical_2_pairs_one_and_two, vertical_2_pairs_two_and_three = vertical_score(v_score)


print("---------SCORING MECHANICS---------")
print()
print("Applies to both horizontal and vertical patterns:")
print("2 PATTERNS(🍒🍒) = $2")
print("3 PATTERNS(🍉🍉🍉) = $3")
print()
print("For horizontal patterns only:")
print("4 PATTERNS(🔔🔔🔔🔔) = $5")
print("5 PATTERNS(💎💎💎💎💎) = $10")
print()
print("---------FAKE SLOT MACHINE---------")
print()

def first_line():
    first_symbol = ""
    for a in slot_line_one:
        while True:
            guess_symbols = random.choice(slot_line_one)
            sys.stdout.write(f"\r      LINE 1:[ {first_symbol + guess_symbols} ]")  # spinning effect.
            sys.stdout.flush()
            time.sleep(.15)  # how fast the animation is
            if guess_symbols == a:
                first_symbol += guess_symbols
                break
    print()


def second_line():
    first_symbol_second_line = ""
    for b in slot_line_two:
        while True:
            guess_symbols_two = random.choice(slot_line_two)
            sys.stdout.write(f"\r      LINE 2:[ {first_symbol_second_line + guess_symbols_two} ]")  # spinning effect.
            sys.stdout.flush()
            time.sleep(.15)  # how fast the animation is
            if guess_symbols_two == b:
                first_symbol_second_line += guess_symbols_two
                break
    print()


def third_line():
    first_symbol_third_line = ""
    for c in slot_line_three:
        while True:
            guess_symbols_three = random.choice(slot_line_three)
            sys.stdout.write(f"\r      LINE 3:[ {first_symbol_third_line + guess_symbols_three} ]")  # spinning effect.
            sys.stdout.flush()
            time.sleep(.15)  # how fast the animation is
            if guess_symbols_three == c:
                first_symbol_third_line += guess_symbols_three
                break
    print()


first_line()
second_line()
third_line()

print()
print("--------------RESULTS--------------")
print("VERTICAL PATTERNS:")
print(f"3 PAIRS - line 1 - 3: {vertical_3_pairs}")
print(f"2 PAIRS - line 1 & 2: {vertical_2_pairs_one_and_two}")
print(f"2 PAIRS - line 2 & 3: {vertical_2_pairs_two_and_three}")
print(f"TOTAL VERTICAL SCORE: ${v_score:.2f}")
print()
print("HORIZONTAL PATTERNS:")
print(f"Line 1 pattern: {h_score_line_one}")
print(f"Line 2 pattern: {h_score_line_two}")
print(f"Line 3 pattern: {h_score_line_three}")
print(f"TOTAL HORIZONTAL SCORE: ${h_score:.2f}")
print()
print(f"YOU WON: ${v_score + h_score:.2f}")
