import docx
from glob import glob


def get_key(d, value):
	for k, v in d.items():
		if v == value:
			return k


alpha = 'а'
letters = {chr(ord(alpha) + i) : 0 for i in range (32)}
general_count = 0

for file in glob(fr"./documents/*.doc*"):
	doc = docx.Document(file)

	for paragraph in doc.paragraphs:
		s = paragraph.text
		s.lower()

		for s1 in s:
			for key in letters:
				if(s1 == key):
					letters[key] += 1
					general_count += 1

letters1 = list(letters.values())
letters1.sort(reverse=True)

print('Статистика по каждой букве:\n')

for letter in letters1:
	#print(get_key(letters, letter), ':', "%.2f" % (letter / general_count * 100), '% (', letter, ')') # old
	print("{:s} : {:^2.4f}% ({:g})".format(get_key(letters, letter), letter / general_count * 100, letter))

print('\nВсего букв: ', general_count)