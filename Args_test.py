#python 3
import argparse

parser = argparse.ArgumentParser(description='Gets some strings and prints'
                                 'either their values or defaults.')
parser.add_argument('-p', dest = 'Position',
                    help = 'Position of the sigining person',
                    default = 'Директор', type = str,
                    required = False)
parser.add_argument('-n', dest = 'Name',
                    default = 'А.В. Кимленко', type = str,
                    help = 'Family Name and initials of the sigining person',
                    required = False)

#this line is necessary to process arguments though
#it should be done in an automatic mode
args = parser.parse_args()
#print(args['Position'])
