# template_engine.py


def replace_template(text, data):

    result = str(text)

    for key, value in data.items():

        placeholder = f"{{{{{key}}}}}"

        result = result.replace(placeholder, str(value))

    return result
