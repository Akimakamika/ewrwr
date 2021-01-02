# -*- coding: utf-8 -*-
"""
Created on Sat Dec 19 14:36:49 2020

@author: Аким
"""

import pandas as pd

# # # # каждый комплект начинается через \n
def from_Akim  ():
    df = pd.read_excel('kucha.xlsx')
    dd = pd.DataFrame(data=df)

    for gg in range(17):
        kucha = dd.iloc[3, gg]
        kucha = kucha.strip()

        # # # # все перечисляются по порядку
        name_from_kucha = kucha.split('\n')
        # Кол-во моделей МИ на странице тупой пдф
        N_N = len(name_from_kucha)
        print('GG', gg + 1)

        # Просто список К-слов, длиною == кол-ву страниц из тупой пдф
        K_word = [i for i in range(17)]
        komissiya = [[K_word[0]] * N_N] * 17
        # все по порядку
        for i in range(N_N):
            komissiya[0][i] = name_from_kucha[i]

        # Создание словаря
        d = {K_word[i]: [''] * N_N for i in range(17)}

        # создание DataFrame
        data_chn = pd.DataFrame(data=d)

        N_modeley = N_N
        # print(komissiya[gg])
        naimenovanie = [''] * N_modeley
        razmer = [' Размер мм.: '] * N_modeley
        razmer_rab = ['Размер рабочей части: '] * N_modeley
        artikul = ['Артикул: '] * N_modeley
        rev_naim = ''

        for i in range(N_modeley):
            name_from_kucha[i] = komissiya[0][i].strip()
            # name_from_kucha[i] = name_from_kucha[i]
            rev_naim = ''.join(reversed(name_from_kucha[i]))

            artikul[i] = rev_naim.partition(' ')[0]
            razmer_rab[i] = (rev_naim.partition(' ')[-1]).partition(' ')[0]
            razmer[i] = ((rev_naim.partition(' ')[-1]).partition(' ')[-1]).partition(' ')[0]

            artikul[i] = ''.join(reversed(artikul[i]))
            razmer_rab[i] = ''.join(reversed(razmer_rab[i]))
            razmer[i] = ''.join(reversed(razmer[i]))

            name_from_kucha[i] = ''.join(reversed(name_from_kucha[i]))
            name_from_kucha[i] = name_from_kucha[i].partition(' ')[-1].partition(' ')[-1].partition(' ')[-1]
            name_from_kucha[i] = ''.join(reversed(name_from_kucha[i]))

            if razmer[i].isdigit():
                razmer[i] = razmer[i]
            else:
                if name_from_kucha[i] == '':
                    name_from_kucha[i] = razmer_rab[i]
                    razmer[i] = ''
                    razmer_rab[i] = ''
                else:
                    name_from_kucha[i] += ' ' + razmer[i]
                    razmer[i] = razmer_rab[i]
                    razmer_rab[i] = ''

# =============================================================================
#             naimenovanie[i] = (name_from_kucha[i] +
#                                ': Размер мм.: ' + razmer[i] + '; ' +
#                                'Размер рабочей части: ' + razmer_rab[i] + '; ' +
#                                'Артикул: ' + artikul[i])
# =============================================================================
            naimenovanie[i] = (name_from_kucha[i] +
                               ' Артикул: ' + artikul[i])
            data_chn.iloc[i, 0] = naimenovanie[i]

            # # # Saving data with indexes
            data_chn.to_excel(f'modified_Ak_V3_{gg}.xlsx')
