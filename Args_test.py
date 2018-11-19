#python 3
import argparse

parser = argparse.ArgumentParser(description='Gets some strings and prints'
                                 'either their values or defaults.')
parser.add_argument('-p', dest = 'Position',
                    help = 'Position of the signing person',
                    default = 'Директор', type = str,
                    required = False)
parser.add_argument('-n', dest = 'Name',
                    default = 'А.В. Кимленко', type = str,
                    help = 'Family Name and initials of the signing person',
                    required = False)
parser.add_argument('-f', dest = 'FileNames',
                    help = 'PDF files to be processed',
                    nargs='+',
                    required = True)

#this line is necessary to process arguments though
#it should be done in an automatic mode
args = parser.parse_args()
for i in args.FileNames:
    print(i)
print(args.Position)
