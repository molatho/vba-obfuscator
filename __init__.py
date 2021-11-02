import argparse
from custom.xorstr import XorStringMutator
import os
import string
import sys
from typing import List
from vba.mutators.strings import StringSplitter

import vba

STRING_MUTATORS = {
    'split': StringSplitter,
    'xor': XorStringMutator
    # Add your custom string mutator here!
}


def createParser():

    parser = argparse.ArgumentParser(prog="vba-obfuscator",
                                     description="Obfuscates VBA code")
    parser.add_argument("input",
                        type=str,
                        help="The input file to parse and obfuscate")
    parser.add_argument("-im",
                        metavar="IGNORE METHODS",
                        dest="imethods",
                        type=str,
                        nargs='*',
                        default=[],
                        help="A list of methods to ignore when obfuscating",
                        action="extend")
    parser.add_argument("-sel",
                        action='store_true',
                        default=False,
                        help="Strip empty lines from code")
    parser.add_argument("-sc",
                        action='store_true',
                        default=False,
                        help="Strip comments from code")
    parser.add_argument("-rall",
                        action='store_true',
                        default=False,
                        help="Rename all identifiers (methods, variables, parameters)")
    parser.add_argument("-rm",
                        action='store_true',
                        default=False,
                        help="Rename methods")
    parser.add_argument("-rv",
                        action='store_true',
                        default=False,
                        help="Rename variables")
    parser.add_argument("-rp",
                        action='store_true',
                        default=False,
                        help="Rename parameters")
    parser.add_argument("-rl",
                        type=int,
                        default=8,
                        help="General length of generated names")
    parser.add_argument("-ra",
                        type=str,
                        default=string.ascii_letters,
                        help=f'General alphabet to use for generated names (default: {string.ascii_letters})')
    parser.add_argument("-rml",
                        type=int,
                        help="Length of generated method names")
    parser.add_argument("-rma",
                        type=str,
                        help=f'Alphabet to use for generated method names (default: {string.ascii_letters})')
    parser.add_argument("-rvl",
                        type=int,
                        help="Length of generated variable names")
    parser.add_argument("-rva",
                        type=str,
                        help=f'Alphabet to use for generated variable names (default: {string.ascii_letters})')
    parser.add_argument("-rpl",
                        type=int,
                        help="Length of generated parameter names")
    parser.add_argument("-rpa",
                        type=str,
                        help=f'Alphabet to use for generated parameter names (default: {string.ascii_letters})')
    parser.add_argument("-strmut",
                        type=str,
                        nargs='*',
                        default=[],
                        choices=list(STRING_MUTATORS.keys()),
                        help="List of string mutator to apply to all strings",
                        action="extend")
    parser.add_argument("-v",
                        action='store_true',
                        default=False,
                        help="Verbose output")
    return parser


def main(args: List[str]):
    parser = createParser()
    args = parser.parse_args(args)

    with open(args.input, "r") as input:
        file = vba.parsing.parse(input, vba.parsing.ParserArguments(
            skipEmptyLines=args.sel,
            stripComments=args.sc,
            verbose=args.v
        ))

    rml = args.rml if args.rml is not None else args.rl
    rpl = args.rpl if args.rpl is not None else args.rl
    rvl = args.rvl if args.rvl is not None else args.rl
    rma = args.rma if args.rma is not None else args.ra
    rpa = args.rpa if args.rpa is not None else args.ra
    rva = args.rva if args.rva is not None else args.ra

    for m in file.methods:
        if m.name in args.imethods:
            continue
        if args.strmut:
            for mutname in args.strmut:
                mut = STRING_MUTATORS.get(mutname, None)
                if mut is None:
                    raise Exception(f'Invalid string mutator "{args.strmut}"')
                for c in m.codeLinesIter:
                    for s in c.parseStrings():
                        mut.process(s)
        if args.rall or args.rm:  # Rename methods
            file.renameMethod(m, name=vba.randomName(rml, rma), verbose=args.v)
        if args.rall or args.rp:  # Rename parameters
            for p in m.parameters:
                m.renameParameter(p, vba.randomName(rpl, rpa), verbose=args.v)
        if args.rall or args.rv:  # Rename variables
            for v in m.variables:
                m.renameVariable(v, vba.randomName(rvl, rva), verbose=args.v)

    _dir, _file = os.path.split(args.input)
    _name, _ext = os.path.splitext(_file)
    _output = os.path.join(_dir, f"{_name}_obf{_ext}")  # path/file_obf.ext

    with open(_output, "w") as output:
        output.writelines(map(lambda l: l.exportLine + "\n", file.dump()))


if __name__ == '__main__':
    main(sys.argv[1:])
