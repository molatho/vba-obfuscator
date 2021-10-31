import vba

with open("sample.vb", "r") as sample:
    file = vba.parse(sample.readlines())

for m in file.methods:
    file.renameMethod(m)
    for v in m.variables:
        m.renameVariable(v)
    for p in m.parameters:
        m.renameParameter(p)

with open("sample_obf.vba", "w") as outp:
    outp.writelines(map(lambda l: l.exportLine + "\n", file.dump()))

print(file)
