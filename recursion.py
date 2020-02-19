# #EXAMPLE RECURSION
# data_dict = {'count': 2,
#             'text': '1',
#             'kids': [{'count': 3,
#                       'text': '1.1',
#                       'kids': [{'count': 1,
#                                 'text': '1.1.1',
#                                 'kids': [{'count':0,
#                                           'text': '1.1.1.1',
#                                           'kids': []}]},
#                                {'count': 0,
#                                 'text': '1.1.2',
#                                 'kids': []},
#                                {'count': 0,
#                                 'text': '1.1.3',
#                                 'kids': []}]},
#                      {'count': 0,
#                       'text': '1.2',
#                       'kids': []}]}
#
#
# def traverse(data):
#     print(' ' * traverse.level + data['text'])
#     for kid in data['kids']:
#         traverse.level += 1
#         traverse(kid)
#         traverse.level -= 1
#
#
# if __name__ == '__main__':
#     traverse.level = 1
#     traverse(data_dict)