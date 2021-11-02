from abc import ABC
from vba.model import String


class StringMutator(ABC):
    def process(st: String):
        raise NotImplementedError()


class StringSplitter(StringMutator):
    def process(st: String):
        # st.tags.add(StringSplitter) # TODO: Strip tags?
        if len(st.value) > 8:
            val = st.value[1:-1]  # Strip quotes
            newStrings = ' & '.join([
                f'"{chunk}"'
                for chunk in StringSplitter.chunkstring(val, 8)  # TODO: Make this customizable
            ])
            st.codeLine.replaceString(st, newStrings)

    # https://stackoverflow.com/a/18854817
    def chunkstring(string, length):
        return (string[0+i:length+i] for i in range(0, len(string), length))
