from dataclasses import dataclass
import re
from typing import Iterable, List
from vba.model import CodeLine, CodeLineType, File, Method, Parameter, ParameterType, Variable
from vba.regex import RE_COMMENT, RE_DECLARATIONS, RE_END_FUNC, RE_METHODS, RE_PARAMETERS, RE_VARIABLES


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
