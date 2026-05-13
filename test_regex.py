import re

def parse_hardware_info(board_text):
    clean_text = board_text.replace(" ", "")
    # Robust regex: Board name, then any chars until parenthesis, then Num_LetterNum
    # Example: G2OH9S01(1_B12)
    match = re.search(r"([A-Z0-9]+).*?\((\d+)_([A-Z])(\d+)\)", clean_text)
    if match:
        return match.groups()
    return None

test_cases = [
    "G 2 O H 9 S 0 1 (1_B12)",
    "G 2 D A P (1_B01)",
    "G 2 W S M D 9 0 1 (1_B04)",
    "O P M 8 (1_B06)",
    "M808SAF3 191.650000THz+-75.0GHz(1_C01)"
]

for tc in test_cases:
    print(f"Testing: '{tc}' -> {parse_hardware_info(tc)}")
