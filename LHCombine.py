import re

def read_and_parse_file(filename):
    with open(filename, 'r') as file:
        content = file.read()

    # Adjusting the regular expression to capture the full text for each entry
    entries = re.findall(r'(#.*?(L\d+|H\d+).*?//)', content, re.DOTALL)
    return {match[1]: match[0] for match in entries}

def combine_light_and_heavy_chains(light_chain_file, heavy_chain_file, output_file):
    light_chains = read_and_parse_file(light_chain_file)
    heavy_chains = read_and_parse_file(heavy_chain_file)

    combined_chains = []
    for i in range(1, max(len(light_chains), len(heavy_chains)) + 1):
        light_key = f'L{i}'
        heavy_key = f'H{i}'
        
        if light_key in light_chains:
            combined_chains.append(light_chains[light_key])
        if heavy_key in heavy_chains:
            combined_chains.append(heavy_chains[heavy_key])

    with open(output_file, 'w') as file:
        file.write('\n'.join(combined_chains))

# Example usage
combine_light_and_heavy_chains('box1light_vj_aa_1-100.txt', 'box1heavy_vdj_aa_1-100.txt', 'combined_chains.txt')
