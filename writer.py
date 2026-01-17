def save(data, filepath):
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(data)
