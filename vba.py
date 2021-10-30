from dataclasses import dataclass
import enum
import re
from typing import Iterator, List
import random
import string

RE_DECLARATIONS = r"^\s*Dim\s+(?P<vars>.*)$"
RE_VARIABLES = r"(?P<name>[\w\d]+)(\sAs\s(?P<type>[\w\d]+(\s+\*\s+\d+)?))?"
RE_FUNCTIONS = r"(^(?P<mod>Private|Public)\s+)?(?P<type>Function|Sub)\s+(?P<name>.*)\((?P<params>(.*)?)\)(\s+As\s+(?P<return>.+))?"
RE_PARAMETERS = r"((?P<tname>\S+) As (?P<type>\w+))|(ByVal (?P<bvname>\S+))|(ByRef (?P<brname>\S+))|((?P<name>[\w\d]+))"
RE_END_FUNC = r"End (Function|Sub)"
RE_COMMENT = r"'(.*)"

NAMES = set()
DEFAULT_NAME_LENGTH = 2


def randomName(length=DEFAULT_NAME_LENGTH, alphabet: str = string.ascii_letters):
    while True:
        name = ''.join(random.choices(alphabet, k=length))
        if name in NAMES:
            continue
        NAMES.add(name)
        return name


@dataclass
class Variable:
    name: str
    newName: str
    type: str

    @property
    def declaration(self) -> str:
        if self.type:
            return f"{self.exportName} As {self.type}"
        else:
            return self.exportName

    @property
    def exportName(self) -> str:
        return self.newName if self.newName else self.name


class ParameterType(enum.Enum):
    Plain = 0,
    Typed = 1,
    ByRef = 2,
    ByVal = 3


@dataclass
class Parameter:
    var: Variable
    ptype: ParameterType

    @property
    def declaration(self) -> str:
        if self.ptype == ParameterType.Plain or self.ptype == ParameterType.Typed:
            return self.var.declaration
        elif self.ptype == ParameterType.ByRef:
            return f"ByRef {self.var.exportName}"
        elif self.ptype == ParameterType.ByVal:
            return f"ByVal {self.var.exportName}"
        else:
            raise Exception("Invalid parameter declaration")


@dataclass
class Method:
    returnType: str
    modifier: str
    name: str
    type: str
    newName: str
    parameters: List[Variable]
    variables: List[Variable]
    lines: List[str]

    @property
    def exportName(self) -> str:
        return self.newName if self.newName else self.name

    @property
    def signature(self) -> str:
        params = ', '.join(map(lambda p: p.declaration, self.parameters))
        return ' '.join(
            filter(lambda e: e is not None,
                   [self.modifier, self.type, f"{self.exportName}({params})", f"As {self.returnType}" if self.returnType else None]
                   )
        )

    def renameVariable(self, var: Variable, name: str = None):
        name = name if name is not None else randomName(DEFAULT_NAME_LENGTH)
        var.newName = name

        # Find variable uses
        pattern = r"(?P<pre>[\W\D])" + var.name + r"(?P<post>[\W\D])?"
        for i, line in enumerate(self.lines[1:-1]):
            newline, subs = re.subn(pattern, r"\g<pre>" + name + r"\g<post>", line)
            if subs > 0:
                print(f'Use of variable "{var.name}", changed from "{line.strip()}" to "{newline.strip()}"')
                self.lines[i+1] = newline

    def renameParameter(self, param: Parameter, name: str = None):
        name = name if name is not None else randomName(DEFAULT_NAME_LENGTH)
        param.var.newName = name

        # Find parameter uses
        pattern = r"(?P<pre>[\W\D])" + param.var.name + r"(?P<post>[\W\D])?"
        for i, line in enumerate(self.lines[1:-1]):
            newline, subs = re.subn(pattern, r"\g<pre>" + name + r"\g<post>", line)
            if subs > 0:
                print(f'Use of parameter "{param.var.name}", changed from "{line.strip()}" to "{newline.strip()}"')
                self.lines[i+1] = newline


@dataclass
class File:
    methods: List[Method]

    def renameMethod(self, meth: Method, name: str = None):
        name = name if name is not None else randomName(DEFAULT_NAME_LENGTH)
        meth.newName = name

        # Find method calls
        pattern = r"\s+" + meth.name + r"\("  # 'method('
        for m in self.methods:
            for i, line in enumerate(m.lines[1:-1]):
                newline, subs = re.subn(pattern, name + "(", line)
                if subs > 0:
                    print(f'Call to {meth.name} in method {m.name}, changed from "{line.strip()}" to "{newline.strip()}"')
                    m.lines[i+1] = newline
        # Find references
        pattern = r"^(?P<indent>\s*)(?P<name>" + meth.name + r")(?P<epilogue>\s+.*)$"  # 'method [= ...|args]
        for m in self.methods:
            for i, line in enumerate(m.lines[1:-1]):
                newline, subs = re.subn(pattern, r"\g<indent>" + name + r"\g<epilogue>", line)
                if subs > 0:
                    print(f'Reference to {meth.name} in method {m.name}, changed from "{line.strip()}" to "{newline.strip()}"')
                    m.lines[i+1] = newline

    def dump(self) -> Iterator[str]:
        for m in self.methods:
            yield m.signature
            for line in m.lines[1:]:
                yield line
            yield ""
            yield ""


def parseParameters(params: str) -> List[Parameter]:
    parameters = []
    for param in re.finditer(RE_PARAMETERS, params):
        if param.group("tname"):
            parameters.append(Parameter(
                var=Variable(name=param.group("tname"),
                             type=param.group("type"),
                             newName=None),
                ptype=ParameterType.Typed
            ))
        elif param.group("bvname"):
            parameters.append(Parameter(
                var=Variable(name=param.group("bvname"),
                             type=None,
                             newName=None),
                ptype=ParameterType.ByVal
            ))
        elif param.group("brname"):
            parameters.append(Parameter(
                var=Variable(name=param.group("brname"),
                             type=None,
                             newName=None),
                ptype=ParameterType.ByRef
            ))
        elif param.group("name"):
            parameters.append(Parameter(
                var=Variable(name=param.group("name"),
                             type=None,
                             newName=None),
                ptype=ParameterType.Plain
            ))
        else:
            raise Exception(f'Failed to parse parameter "{param.string}"')

    return parameters


def parseVariables(variables: str) -> List[Variable]:
    vars = [
        Variable(name=var.group("name"),
                 type=var.group("type"),
                 newName=None)
        for var in re.finditer(RE_VARIABLES, variables)
    ]
    types = {var.type for var in vars if var.type != None}
    if len(types) > 1:
        raise Exception(f'Found {len(types)} ({types}) in "{variables}"')
    if len(types) == 1 and len(vars) > 1:
        targetType = types.pop()
        for var in vars:
            var.type = targetType
    return vars


def parse(lines: List[str]) -> File:
    meth: Method = None
    file: File = File([])

    for i, line in enumerate(lines):
        # Strip comments
        line = re.sub(RE_COMMENT, "", line).rstrip()
        #if len(line.strip()) == 0:
        #    continue

        # Method definition?
        match = re.match(RE_FUNCTIONS, line)
        if match:
            if meth:
                raise Exception("Defined method before terminating method")

            meth = Method(
                name=match.group("name"),
                type=match.group("type"),
                newName=None,
                returnType=match.group("return"),
                modifier=match.group("mod"),
                parameters=[],
                variables=[],
                lines=[])

            # Parse parameters (if any)
            params = match.group("params")
            if params:
                meth.parameters = parseParameters(params)

            file.methods.append(meth)
            meth.lines.append(line)
            continue

        # Method termination?
        match = re.match(RE_END_FUNC, line)
        if match:
            if not meth:
                raise Exception("Terminated method before defining method")
            meth.lines.append(line)
            meth = None
            continue

        # Variable declaration?
        match = re.match(RE_DECLARATIONS, line)
        if match:
            meth.lines.append(line)
            meth.variables.extend(parseVariables(match.group("vars")))
            continue

        if meth:
            meth.lines.append(line)
        print(f'[{(i+1)}] Unmatched line: "{line}"')

    return file
