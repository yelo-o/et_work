from pdf2image import convert_from_path

page = convert_from_path('paper.pdf')


page.save('paper.jpg', 'JPEG')

