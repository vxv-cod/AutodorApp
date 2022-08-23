def GO(string):
    # legend = {
    #     'а': 'a',
    #     'б': 'b',
    #     'в': 'v',
    #     'г': 'g',
    #     'д': 'd',
    #     'е': 'e',
    #     'ё': 'yo',
    #     'ж': 'zh',
    #     'з': 'z',
    #     'и': 'i',
    #     'й': 'y',
    #     'к': 'k',
    #     'л': 'l',
    #     'м': 'm',
    #     'н': 'n',
    #     'о': 'o',
    #     'п': 'p',
    #     'р': 'r',
    #     'с': 's',
    #     'т': 't',
    #     'у': 'u',
    #     'ф': 'f',
    #     'х': 'h',
    #     'ц': 'ts',
    #     'ч': 'ch',
    #     'ш': 'sh',
    #     'щ': 'shch',
    #     'ъ': 'y',
    #     'ы': 'y',
    #     'ь': "'",
    #     'э': 'e',
    #     'ю': 'yu',
    #     'я': 'ya',
    #     'А': 'A',
    #     'Б': 'B',
    #     'В': 'V',
    #     'Г': 'G',
    #     'Д': 'D',
    #     'Е': 'E',
    #     'Ё': 'Yo',
    #     'Ж': 'Zh',
    #     'З': 'Z',
    #     'И': 'I',
    #     'Й': 'Y',
    #     'К': 'K',
    #     'Л': 'L',
    #     'М': 'M',
    #     'Н': 'N',
    #     'О': 'O',
    #     'П': 'P',
    #     'Р': 'R',
    #     'С': 'S',
    #     'Т': 'T',
    #     'У': 'U',
    #     'Ф': 'F',
    #     'Х': 'H',
    #     'Ц': 'Ts',
    #     'Ч': 'Ch',
    #     'Ш': 'Sh',
    #     'Щ': 'Shch',
    #     'Ъ': 'Y',
    #     'Ы': 'Y',
    #     'Ь': "'",
    #     'Э': 'E',
    #     'Ю': 'Yu',
    #     'Я': 'Ya',
    # }

    legend = {
        'а': 'a',
        'б': 'b',
        'в': 'v',
        'г': 'g',
        'д': 'd',
        'е': 'e',
        'ё': 'yo',
        'ж': 'zh',
        'з': 'z',
        'и': 'i',
        'й': 'y',
        'к': 'k',
        'л': 'l',
        'м': 'm',
        'н': 'n',
        'о': 'o',
        'п': 'p',
        'р': 'r',
        'с': 's',
        'т': 't',
        'у': 'u',
        'ф': 'f',
        'х': 'h',
        'ц': 'ts',
        'ч': 'ch',
        'ш': 'sh',
        'щ': 'shch',
        'ъ': 'y',
        'ы': 'y',
        'ь': "'",
        'э': 'e',
        'ю': 'yu',
        'я': 'ya',
        'А': 'A',
        'Б': 'B',
        'В': 'V',
        'Г': 'G',
        'Д': 'D',
        'Е': 'E',
        'Ё': 'YO',
        'Ж': 'ZH',
        'З': 'Z',
        'И': 'I',
        'Й': 'Y',
        'К': 'K',
        'Л': 'L',
        'М': 'M',
        'Н': 'N',
        'О': 'O',
        'П': 'P',
        'Р': 'R',
        'С': 'S',
        'Т': 'T',
        'У': 'U',
        'Ф': 'F',
        'Х': 'H',
        'Ц': 'TS',
        'Ч': 'CH',
        'Ш': 'SS',
        'Щ': 'SHCH',
        'Ъ': 'Y',
        'Ы': 'Y',
        'Ь': "'",
        'Э': 'E',
        'Ю': 'YU',
        'Я': 'YA',
    }
    new_string = ""
    for s in string:
        if s in legend:
            new_string += legend[s]
        elif s == " " or s == "/" or s == ".":
            new_string += "_"
        else:
            new_string += s

    return new_string


if __name__ == "__main__":
    print(GO(input("Введите строку: ")))