def ordinal_formatter(i):
    suffixes = ["th", "st", "nd", "rd", "th", "th", "th", "th", "th", "th"]
    if (i // 10) % 10 == 1:
        return str(i) + 'th'
    else:
        return str(i) + suffixes[i % 10]


def ordinal(n, size):
    return [ordinal_formatter(place) for place in range(1, size + 1) if place != n]

print(ordinal(5, 37))
