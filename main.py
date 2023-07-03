from docx2txt import process
from pandas import DataFrame


def extract_data(path):
    with open(path, 'rb') as file:
        my_text = process(file).split('\n\n')
        subject, code, topics, subtopics = '', '', [], []
        start_topics, end_topics = 0, 0
        rome_numbers = ['I.', 'II.', 'III.', 'IV.', 'V.', 'VI.', 'VII.', 'VIII.', 'IX.', 'X.', 'XI.', 'XII.', 'XIII.', 'XIV.', 'XV.']

        for i, line in enumerate(my_text):
            if 'Programa de estudios de la asignatura' in line:
                subject = my_text[i + 1]
            elif 'Clave' in line:
                code = my_text[i + 4]
            elif 'Contenido Temático' in line:
                start_topics = i + 1
            elif 'Estrategias didácticas' in line:
                end_topics = i
                break

        topic_data = my_text[start_topics:end_topics]
        objectives = False
        for i, line in enumerate(topic_data):
            if any(number in line for number in rome_numbers):
                topics.append(line)
                if objectives:
                    subtopics.append(data)
                    objectives = False
            elif 'Objetivo' in line:
                objectives = True
                data = []
            elif '\t\t' in line and objectives:
                data.append(line[2:])
        subtopics.append(data)
        return subject, code, topics, subtopics


def extract_path(origin_path):
    path = ''
    for label in origin_path.split('/')[:-1]:
        path += label + '/'
    return path


def save_as_excel(subject, code, topics, subtopics, path, name):
    data = {'Clave Plan de Estudios': ['' for topic in subtopics for _ in topic],
            'Clave Materia': [code for topic in subtopics for _ in topic],
            'Materia': [subject for topic in subtopics for _ in topic],
            'Tema': [topic for i, topic in enumerate(topics) for _ in range(len(subtopics[i]))],
            'Sub Tema': [subtopic for topic in subtopics for subtopic in topic]
            }
    df = DataFrame.from_dict(data)
    df.to_excel(path + '/' + name + '.xlsx', index=False)


if __name__ == '__main__':
    pdf = input('Ingresa la ruta del documento:\n')
    s, c, t, st = extract_data(pdf)
    p = '/'.join(pdf.split('/')[:-1])
    save_as_excel(s, c, t, st, p, s.upper())
