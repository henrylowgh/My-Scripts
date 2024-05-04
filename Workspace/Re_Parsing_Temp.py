import re

# Search for regular expression (a.k.a. partial string subset) that matches [PREFIX]-L1/[PREFIX]-H1 formatting
def parse_description(description):
    # Regex updated to optionally ignore a leading '>' and capture the rest as before
    match = re.match(r'^>?(.+)-(H|L)(\d+)', description)
    if match:
        prefix = match.group(1)  # This will include everything up to the hyphen before "H" or "L"
        chain_type = match.group(2)
        number = match.group(3)
        return f"{prefix}-{chain_type}{number}"
    return None

# Example use
description1 = '>TDM-1-H23-mIgGR1_C04.ab1_(reversed)_-_VDJ-REGION_translation'
description2 = 'TDM-1-H23-mIgGR1_C04.ab1_(reversed)_-_VDJ-REGION_translation'
parsed_description1 = parse_description(description1)
parsed_description2 = parse_description(description2)
print(parsed_description1)  # Should print "TDM-1-H23"
print(parsed_description2)  # Should also print "TDM-1-H23"