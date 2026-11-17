with open('frontend/index.html', 'r', encoding='utf-8') as f:
    c = f.read()

idx = c.find('function renderBatchResults')
print(c[idx+6000:idx+7800])
