from dataclasses import dataclass, replace
import enum
import re
from typing import Iterator, List
import random
import string

RE_DECLARATIONS = r"^\s*Dim\s+(?P<vars>.*)$"  # 'Dim ...'
RE_VARIABLES = r"(?P<name>[\w\d]+)(\sAs\s(?P<type>[\w\d]+(\s+\*\s+\d+)?))?"  # 'var [As String]'
RE_METHODS = r"(^(?P<mod>Private|Public)\s+)?(?P<type>Function|Sub)\s+(?P<name>.*)\((?P<params>(.*)?)\)(\s+As\s+(?P<return>.+))?"  # '[Private] Function func([params]) [As String]
RE_METHOD_CALLS = r"(?P<pre>(^\s*|\s+))%NAME%\("  # 'method('
RE_METHOD_CALLS_SUB = r"\g<pre>%NAME%("
RE_METHOD_REF = r"^(?P<pre>\s*)%NAME%(?P<post>\s+.*)$"  # 'method [= ...|args]'
RE_METHOD_REF_SUB = r"\g<pre>%NAME%\g<post>"
RE_PARAMETERS = r"((?P<tname>\S+) As (?P<type>\w+))|(ByVal (?P<bvname>\S+))|(ByRef (?P<brname>\S+))|((?P<name>[\w\d]+))"  # 'var As String'
RE_END_FUNC = r"End (Function|Sub)"  # 'End Function
RE_COMMENT = r"'(.*)"
RE_IDENTIFIER_USE = r"(?P<pre>[\W\D])%NAME%(?P<post>[\W\D])?"  # Matches %NAME% when there's no [a-zA-Z0-9] prepended or appended (=> exact identifier only), but allows ',' or '(' etc
RE_IDENTIFIER_SUB = r"\g<pre>%NAME%\g<post>"

NAMES = set()
DEFAULT_NAME_LENGTH = 8


def randomName(length=DEFAULT_NAME_LENGTH, alphabet: str = string.ascii_letters):
    while True:
        name = ''.join(random.choices(alphabet, k=length))
        if name in NAMES:
            continue
        NAMES.add(name)
        return name


@dataclass
class CodeLine:
    line: str
    newLine: str
    number: int

    @property
    def exportLine(self) -> str:
        return self.newLine if self.newLine else self.line

    @exportLine.setter
    def exportLine(self, newLine: str):
        self.newLine = newLine

    @property
    def isEmpty(self) -> str:
        return len(self.exportLine.strip()) == 0


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
    codeLines: List[CodeLine]

    @property
    def exportName(self) -> str:
        return self.newName if self.newName else self.name

    @exportName.setter
    def exportName(self, newName):
        self.newName = newName
        self.codeLines[0].exportLine = self.signature

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

        replaceIdentifier(var.name, var.newName, self.codeLines[1:-1])

    def renameParameter(self, param: Parameter, name: str = None):
        name = name if name is not None else randomName(DEFAULT_NAME_LENGTH)
        param.var.newName = name

        replaceIdentifier(param.var.name, param.var.newName, self.codeLines[1:-1])


def replaceIdentifier(oldName: str, newName: str, lines: List[CodeLine]):
    pattern = RE_IDENTIFIER_USE.replace("%NAME%", oldName)
    for codeLine in lines:
        newline, subs = re.subn(pattern,
                                RE_IDENTIFIER_SUB.replace("%NAME%", newName),
                                codeLine.exportLine)
        if subs > 0:
            print(f'[{codeLine.number}] Reference to identifier "{oldName}" changed from "{codeLine.exportLine.strip()}" to "{newline.strip()}"')
            codeLine.exportLine = newline


@dataclass
class File:
    methods: List[Method]
    lines: List[CodeLine]

    def renameMethod(self, meth: Method, name: str = None):
        name = name if name is not None else randomName(DEFAULT_NAME_LENGTH)
        meth.exportName = name

        # Find method calls
        pattern = RE_METHOD_CALLS.replace("%NAME%", meth.name)
        # r"\s+" + meth.name + r"\("  # 'method('
        for m in self.methods:
            for i, codeLine in enumerate(m.codeLines[1:-1]):
                newline, subs = re.subn(pattern,
                                        RE_METHOD_CALLS_SUB.replace("%NAME%", name),
                                        codeLine.exportLine)
                if subs > 0:
                    print(f'Call to {meth.name} in method {m.name}, changed from "{codeLine.exportLine.strip()}" to "{newline.strip()}"')
                    codeLine.exportLine = newline
        # Find references
        pattern = RE_METHOD_REF.replace("%NAME%", meth.name)
        for m in self.methods:
            for i, codeLine in enumerate(m.codeLines[1:-1]):
                newline, subs = re.subn(pattern,
                                        RE_METHOD_REF_SUB.replace("%NAME%", name),
                                        codeLine.exportLine)
                if subs > 0:
                    print(f'Reference to {meth.name} in method {m.name}, changed from "{codeLine.exportLine.strip()}" to "{newline.strip()}"')
                    codeLine.exportLine = newline

    def dump(self) -> Iterator[str]:
        return self.lines
        # for m in self.methods:
        #     yield m.signature
        #     for codeLine in m.codeLines[1:]:
        #         yield codeLine.exportLine
        #     yield ""
        #     yield ""


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
    file: File = File(methods=[], lines=[])

    for i, line in enumerate(lines):
        # Strip comments
        line = re.sub(RE_COMMENT, "", line).rstrip()
        codeLine = CodeLine(line=line, newLine=None, number=(i+1))
        file.lines.append(codeLine)
        # if len(line.strip()) == 0:
        #    continue

        # Method definition?
        match = re.match(RE_METHODS, codeLine.line)
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
                codeLines=[])

            # Parse parameters (if any)
            params = match.group("params")
            if params:
                meth.parameters = parseParameters(params)

            file.methods.append(meth)
            meth.codeLines.append(codeLine)
            continue

        # Method termination?
        match = re.match(RE_END_FUNC, line)
        if match:
            if not meth:
                raise Exception("Terminated method before defining method")
            meth.codeLines.append(codeLine)
            meth = None
            continue

        # Variable declaration?
        match = re.match(RE_DECLARATIONS, line)
        if match:
            meth.codeLines.append(codeLine)
            meth.variables.extend(parseVariables(match.group("vars")))
            continue

        if meth:
            meth.codeLines.append(codeLine)
        print(f'[{(i+1)}] Unmatched line: "{line}"')

    return file
