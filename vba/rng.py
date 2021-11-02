import random
import string
from typing import Dict, Set, Tuple


NAMES: Dict[Tuple[int, str], Set[str]] = {}  # Saves all generated names per length/alphabet combination
DEFAULT_NAME_LENGTH = 8  # TODO: Remove


def randomName(length=DEFAULT_NAME_LENGTH, alphabet: str = string.ascii_letters):
    names = NAMES.get((length, alphabet), None)
    if names is None:
        names = NAMES[(length, alphabet)] = set()

    MAX = pow(len(alphabet), length)
    if len(names) >= MAX:
        raise Exception(f'Exceeded maximum of {MAX} random names with {length} elements of alphabet "{alphabet}"')

    while True:
        name = ''.join(random.choices(alphabet, k=length))
        if name in NAMES:
            continue
        names.add(name)
        return name
