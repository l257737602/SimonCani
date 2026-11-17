import re

def parse_points(input_str):
    # Use regex to find all tuples in the input string
    pattern = r'\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)'
    points = re.findall(pattern, input_str)
    return [(int(x), int(y)) for x, y in points]

def average_point(shifts, *inputs):
    total_x = 0
    total_y = 0
    count = 0

    for input_str in inputs:
        name, coords_str = [string.strip() for string in input_str.split('=')]
        shift = shifts.get(name, (0, 0)) # Default shift is (0, 0)
        points = parse_points(coords_str)
        
        for x, y in points:
            total_x += x + shift[0]
            total_y += y + shift[1]
            count += 1

    if count == 0:
        return None  # No points to average

    return (total_x / count, total_y / count)

# Define the shifts
marks_above = {
    'tilde'   : (228, 578),
    'dot'     : (244, 578),
    'ring'    : (231, 578),
    'dieresis': (224, 578),
    'linevert': (231, 578),
    'breve'   : (231, 578),
    'comma'   : (231, 578),
}

reference_points_above = {
    'tilde'   : (361, 597), # extreme up right
    'dot'     : (244, 616), # extreme up
    'dieresis': (319, 578), # up, right
    'ring'    : (231, 637), # up, inside the circle
    'linevert': (208, 665), # up, left
    'breve'   : (231, 607), # up, extreme point
    'comma'   : (213, 681), # up, extreme point
}

shifts_above = {key: (
        marks_above[key][0] - reference_points_above[key][0],
        marks_above[key][1] - reference_points_above[key][1]
    ) for key in reference_points_above}

# Call the function
average_above = average_point(
    shifts_above,
    "comma = (197, -62), (206, -62)",
    "tilde = (338, -77), (350, -83)",
    "breve = (220, -216)",
    "ring = (232, -125)",
)
print("average Point (above):", average_above)

# Define the shifts
marks_below = {
    'dot'     : (278, -184),
    'dieresis': (280, -184),
    'ring'    : (271, -184),
    'comma'   : (269, -184),
    'linevert': (272, -184),
    'breve'   : (270, -184),
    'tilde'   : (273, -200),
}

reference_points_below = {
    'dot'     : (279, -109), # extreme up
    'dieresis': (376, -101), # up, right
    'ring'    : (271, -125), # up, inside the circle
    'comma'   : (245, -62), # up, extreme point
    'linevert': (248, -65), # up, left
    'breve'   : (271, -216), # down, extreme point
    'tilde'   : (407, -100), # extreme up right
}

shifts_below = {key: (
        marks_below[key][0] - reference_points_below[key][0],
        marks_below[key][1] - reference_points_below[key][1]
    ) for key in reference_points_below}

# Call the function
average_below = average_point(
    shifts_below,
    "comma = (197, -62), (206, -62)",
    "tilde = (338, -77), (350, -83)",
    "breve = (220, -216)",
    "ring = (232, -125)",
)
print("average Point (below):", average_below)
