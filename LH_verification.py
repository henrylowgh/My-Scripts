import re

def verify_order(filename):
    with open(filename, 'r') as file:
        content = file.read()

    # Extracting the identifiers (L1, H1, etc.) from the entries
    identifiers = re.findall(r'(#.*?(L\d+|H\d+))', content)
    identifiers = [identifier[1] for identifier in identifiers]

    # Verifying the L -> H -> L -> H order
    for i in range(len(identifiers) - 1):
        current_id, next_id = identifiers[i], identifiers[i + 1]
        current_type, next_type = current_id[0], next_id[0]

        if current_type == next_type:
            return False, f"Order mismatch at {current_id} and {next_id}"

    return True, "Order is correct"

# Example usage
filename = 'Workspace/box1combined_aa_finalized.txt' # Replace with your actual file name
is_correct, message = verify_order(filename)
print(message)
