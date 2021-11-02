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
