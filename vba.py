from dataclasses import dataclass, replace
import enum
import re
from typing import Dict, Iterable, Iterator, List, Set, Tuple
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
RE_STRINGS = r"\"[^\"]*\""  # Matches everything between two double-quotes

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


@dataclass
class String:
    value: str
    startIdx: int
    endIdx: int
    codeLine: 'CodeLine'
    newString: str = None

    @property
    def exportString(self) -> str:
        return self.newString if self.newString is not None else self.value

    @exportString.setter
    def exportString(self, newString: str):
        self.newString = newString


class CodeLineType(enum.Enum):
    Default = 0,
    MethodStart = 1,
    MethodEnd = 2


class CodeLine:
    def __init__(self, line: str, number: int = -1, method: 'Method' = None, lineType: CodeLineType = CodeLineType.Default):
        self.line: str = line
        self.number: int = number
        self.method: Method = method
        self.lineType: CodeLineType = lineType
        self._newLine: str = None
        self._prev: CodeLine = None
        self._next: CodeLine = None
        self.strings: List[String] = self.getStrings()

    @property
    def next(self) -> 'CodeLine':
        return self._next

    @next.setter
    def next(self, next: 'CodeLine'):
        if next is self._next:
            return

        old = self._next
        self._next = next

        if old is not None:
            old.prev = None
        if next is not None:
            next.prev = self

    @property
    def isFirst(self) -> bool:
        return self._prev is None

    @property
    def isLast(self) -> bool:
        return self._next is None

    @property
    def first(self) -> 'CodeLine':
        return self if self.isFirst else self._prev

    @property
    def last(self) -> 'CodeLine':
        return self if self.isLast else self._next

    @property
    def prev(self) -> 'CodeLine':
        return self._prev

    @prev.setter
    def prev(self, prev: 'CodeLine'):
        if prev is self._prev:
            return

        old = self._prev
        self._prev = prev

        if old is not None:
            old.next = None
        if prev is not None:
            prev.next = self

    def getStrings(self) -> List[String]:
        return [
            String(value=match.string[match.start():match.end()],
                   startIdx=match.start(),
                   endIdx=match.end(),
                   codeLine=self)
            for match in re.finditer(RE_STRINGS, self.line)
        ]

    @property
    def exportLine(self) -> str:
        return self._newLine if self._newLine else self.line

    @exportLine.setter
    def exportLine(self, newLine: str):
        self._newLine = newLine

    @property
    def isEmpty(self) -> str:
        return len(self.exportLine.strip()) == 0

    def replaceWith(self, newLines: List['CodeLine']):
        _next = self._next
        _prev = self._prev

        if _next is not None:
            _next.insertBefore(newLines)
        elif _prev is not None:
            _prev.insertAfter(newLines)
        else:
            raise Exception("Can't replace line with new lines: no preceeding/succeeding lines to link to")

    def remove(self) -> 'CodeLine':
        if self._next is not None:
            self._next.prev = self._prev
        elif self._prev is not None:
            self._prev.next = self._next
        else:
            raise Exception("Can't remove line: no preceeding/succeeding lines to unlink from")
        return self

    def insertAfter(self, newLines: List['CodeLine']) -> 'CodeLine':
        _next = self._next
        meth = _next.method if next is not None else None
        _line = self
        for line in newLines:
            _line.next = line
            _line = line
            _line.method = meth
        _line.next = _next
        return _line

    def insertBefore(self, newLines: List['CodeLine']) -> 'CodeLine':
        _prev = self._prev
        meth = _prev.method if next is not None else None
        _line = self
        for line in newLines:
            _line.prev = line
            _line = line
            _line.method = meth
        _line.prev = _prev
        return _line

    def replaceString(self, st: String, repl: str):
        following = [string for string in self.strings if string.startIdx > st.endIdx]
        diff = len(repl) - len(st.value)

        # Update line contents
        self.exportLine = self.exportLine[:st.startIdx] + repl + self.exportLine[st.endIdx:]
        st.endIdx = st.startIdx + len(repl)
        st.newString = repl

        for _st in following:  # Update indices of following strings
            _st.startIdx += diff
            _st.endIdx += diff

    def __repr__(self) -> str:
        return self.exportLine


@dataclass
class Variable:
    name: str
    type: str = None
    newName: str = None

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


class Method:
    def __init__(self, file: 'File', returnType: str, modifier: str, name: str, type: str):
        self.returnType: str = returnType
        self.modifier: str = modifier
        self.name: str = name
        self.type: str = type
        self._file: 'File' = file
        self.parameters: List[Variable] = []
        self.variables: List[Variable] = []

    @property
    def codeLinesIter(self) -> Iterator[CodeLine]:
        return (line for line in self._file.originalLines if line.method is self)

    @property
    def codeLines(self) -> List[CodeLine]:
        return list(self.codeLinesIter)

    @property
    def exportName(self) -> str:
        return self.newName if self.newName else self.name

    @property
    def firstLine(self) -> CodeLine:
        return next(self.codeLinesIter)

    @property
    def lastLine(self) -> CodeLine:
        iter = next(self.codeLinesIter)
        last = next(iter)
        for last in iter:
            pass
        return last

    @exportName.setter
    def exportName(self, newName):
        self.newName = newName
        self.firstLine.exportLine = self.signature

    @property
    def signature(self) -> str:
        params = ', '.join(map(lambda p: p.declaration, self.parameters))
        return ' '.join(
            filter(lambda e: e is not None,
                   [self.modifier, self.type, f"{self.exportName}({params})", f"As {self.returnType}" if self.returnType else None]
                   )
        )

    def renameVariable(self, var: Variable, name: str = None, verbose: bool = False):
        name = name if name is not None else randomName(DEFAULT_NAME_LENGTH)
        var.newName = name

        replaceIdentifier(var.name, var.newName, self.codeLines[1:-1])

    def renameParameter(self, param: Parameter, name: str = None, verbose: bool = False):
        name = name if name is not None else randomName(DEFAULT_NAME_LENGTH)
        param.var.newName = name

        replaceIdentifier(param.var.name, param.var.newName, self.codeLines[1:-1])


def replaceIdentifier(oldName: str, newName: str, lines: List[CodeLine], verbose: bool = False):
    pattern = RE_IDENTIFIER_USE.replace("%NAME%", oldName)
    for codeLine in lines:
        newline, subs = re.subn(pattern,
                                RE_IDENTIFIER_SUB.replace("%NAME%", newName),
                                codeLine.exportLine)
        if subs > 0:
            if verbose:
                print(f'[{codeLine.number}] Reference to identifier "{oldName}" changed from "{codeLine.exportLine.strip()}" to "{newline.strip()}"')
            codeLine.exportLine = newline


@dataclass
class File:
    def __init__(self):
        self.methods: List[Method] = []
        self.originalLines: List[CodeLine] = []

    def renameMethod(self, meth: Method, name: str = None, verbose: bool = False):
        name = name if name is not None else randomName(DEFAULT_NAME_LENGTH)
        meth.exportName = name

        # Find method calls
        pattern = RE_METHOD_CALLS.replace("%NAME%", meth.name)
        for m in self.methods:
            for i, codeLine in enumerate(m.codeLines[1:-1]):
                newline, subs = re.subn(pattern,
                                        RE_METHOD_CALLS_SUB.replace("%NAME%", name),
                                        codeLine.exportLine)
                if subs > 0:
                    if verbose:
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
                    if verbose:
                        print(f'Reference to {meth.name} in method {m.name}, changed from "{codeLine.exportLine.strip()}" to "{newline.strip()}"')
                    codeLine.exportLine = newline

    def dump(self) -> Iterator[str]:
        line = self.originalLines[0]
        while line:
            yield line
            line = line.next


def parseParameters(params: str) -> List[Parameter]:
    parameters = []
    for param in re.finditer(RE_PARAMETERS, params):
        if param.group("tname"):
            parameters.append(Parameter(
                var=Variable(name=param.group("tname"),
                             type=param.group("type")),
                ptype=ParameterType.Typed
            ))
        elif param.group("bvname"):
            parameters.append(Parameter(
                var=Variable(name=param.group("bvname")),
                ptype=ParameterType.ByVal
            ))
        elif param.group("brname"):
            parameters.append(Parameter(
                var=Variable(name=param.group("brname")),
                ptype=ParameterType.ByRef
            ))
        elif param.group("name"):
            parameters.append(Parameter(
                var=Variable(name=param.group("name")),
                ptype=ParameterType.Plain
            ))
        else:
            raise Exception(f'Failed to parse parameter "{param.string}"')

    return parameters


def parseVariables(variables: str) -> List[Variable]:
    vars = [
        Variable(name=var.group("name"),
                 type=var.group("type"))
        for var in re.finditer(RE_VARIABLES, variables)
    ]
    # Sanity-check:
    # If we found more than one variable in the search string, make sure they are of the same type.
    # If there are multiple types, that's invalid syntax (last variable dictates type for all).
    # If there is one type annotation, apply this to all variables.
    types = {var.type for var in vars if var.type != None}
    if len(types) > 1:
        raise Exception(f'Found {len(types)} ({types}) in "{variables}"')
    if len(types) == 1 and len(vars) > 1:
        targetType = types.pop()
        for var in vars:
            var.type = targetType
    return vars


@dataclass
class ParserArguments:
    skipEmptyLines: bool = True
    stripComments: bool = True
    verbose: bool = True


def parse(lines: Iterable[str], parserArgs: ParserArguments = None) -> File:
    parserArgs = parserArgs if parserArgs is not None else ParserArguments()
    meth: Method = None
    file: File = File()
    last: CodeLine = None

    for i, line in enumerate(lines):
        if parserArgs.stripComments:
            line = re.sub(RE_COMMENT, "", line)
        line = line.rstrip()

        codeLine = CodeLine(line=line, number=(i+1))

        if parserArgs.skipEmptyLines and len(line.strip()) == 0:
            continue

        codeLine.prev = last
        last = codeLine
        file.originalLines.append(codeLine)

        # Method definition?
        match = re.match(RE_METHODS, codeLine.line)
        if match:
            if meth:
                raise Exception("Defined method before terminating method")

            codeLine.lineType = CodeLineType.MethodStart
            meth = Method(
                file=file,
                name=match.group("name"),
                type=match.group("type"),
                returnType=match.group("return"),
                modifier=match.group("mod"))

            # Parse parameters (if any)
            params = match.group("params")
            if params:
                meth.parameters = parseParameters(params)

            file.methods.append(meth)
            # meth.codeLines.append(codeLine)
            codeLine.method = meth
            continue

        # Method termination?
        match = re.match(RE_END_FUNC, line)
        if match:
            if not meth:
                raise Exception("Terminated method before defining method")

            codeLine.lineType = CodeLineType.MethodEnd
            codeLine.method = meth
            meth = None
            continue

        # Variable declaration?
        match = re.match(RE_DECLARATIONS, line)
        if match:
            if meth:
                meth.variables.extend(parseVariables(match.group("vars")))
                codeLine.method = meth
            continue

        if meth:
            codeLine.method = meth

        if parserArgs.verbose:
            print(f'[{(i+1)}] Unmatched line: "{line}"')

    return file
