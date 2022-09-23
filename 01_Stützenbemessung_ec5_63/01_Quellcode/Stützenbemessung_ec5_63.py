
import xlwings as xw
from math import ceil, pi, sqrt
import os
from pyautocad import Autocad, APoint, aDouble, ACAD


# Stützenbemessungsprogramm: Funktion - Ersatzstabverfahren
def ec5_63_esv(Lagerung, güte, L, b, h, N_ed, M_yd, M_zd, theta, f_c0k, f_myk, f_mzk, E, G_05, k_mod, gamma):

    # Vorbereitung
    # Listen
    L_M = [M_yd, M_zd]
    L_bh = [[b, h], [h, b]]
    L_fmk = [f_myk, f_mzk]
    L_k_crit = [1, 1]
    L_lamb = []
    L_lamb_rel = []
    L_ky = []
    L_kc = []
    L_sigma_md = []
    L_eta = []
    L_nw = []
    L_w = []
    L_I = []

    # Beiwert km
    if M_yd == 0 or M_zd == 0:
        L_km = [[1, 1], [1, 1]]
    else:
        L_km = [[1, 0.7], [0.7, 1]]

    # if-Abfrage zur Bestimmung von beta
    if Lagerung == 'Pendelstütze':
        beta = 1
    elif Lagerung == 'Kragstütze':
        beta = 2
    elif Lagerung in ['Eingespannt (o)', 'Eingespannt (u)']:
        beta = 0.7
    else:
        beta = 0.5

    # Berechnung
    # Knicklänge
    l_ef = L*beta
    e_0 = 0

    # Bemessungswerte der Festigkeit
    xi = k_mod/gamma
    f_c0d = f_c0k*xi  # $N/mm^2$ - Bemessungswert der Druckfestigkeit
    f_myd = f_myk*xi  # $N/mm^2$ - Bemessungswert der Biegefestigkeit
    f_mzd = f_mzk*xi  # $N/mm^2$ - Bemessungswert der Biegefestigkeit

    # Schleife für Momente um y- und z-Achse
    for n, M in enumerate(L_M):

        # Querschnittsparameter
        A = b*h  # $m^2$
        I = (L_bh[n][0]*L_bh[n][1]**3)/12  # $m^4$ - FTM
        w = (L_bh[n][0]*L_bh[n][1]**2)/6  # $m^3$ - Widerstandsmoment
        i = L_bh[n][1]/sqrt(12)  # $m$ - polares Trägheitsmoment
        L_w.append(w)
        L_I.append(I)

        # Knicken
        lamb = (beta*l_ef)/i  # Schlankheitsgrad
        lamb_rel = lamb/pi*sqrt(f_c0k/E)  # bezogene Schlankheit
        k_y = 0.5*(1+0.1*(lamb_rel-0.3)+lamb_rel**2)  # Beiwert
        k_c = 1/(k_y+sqrt(k_y**2-lamb_rel**2))  # Knickbeiwert

        # Spannungen
        sigma_cd = N_ed/A  # $kN/m^2$ - Druckspannung
        sigma_md = M/w  # $kN/m^2$ - Biegespannung

        # Anhängen an Listen
        L_lamb.append(round(lamb, 2))
        L_lamb_rel.append(round(lamb_rel, 2))
        L_ky.append(round(k_y, 2))
        L_kc.append(round(k_c, 2))
        L_sigma_md.append(round(sigma_md, 2))

    # Kippen um starke Achse
    # Ermittlung des maßgebenden Widerstandsmoments
    index = L_w.index(max(L_w))
    w = L_w[index]
    b = L_bh[index][0]
    h = L_bh[index][1]
    f_mk = L_fmk[index]

    lamb_relm = sqrt(l_ef/(pi*b**2)) * \
        sqrt(f_mk/sqrt(E*G_05))  # bez. Schlankheitsgrad

    # Kippbeiwert
    if lamb_relm <= 0.75:
        k_crit = 1
    elif lamb_relm > 0.75 and lamb_relm < 1.4:
        k_crit = 1.56-0.75*lamb_relm
    elif lamb_relm > 1.4:
        k_crit = 1/lamb_relm**2

    if index == 0:
        L_k_crit[0] = k_crit
        L_k_crit[1] = 1

    elif index == 1:
        L_k_crit[0] = 1
        L_k_crit[1] = k_crit

    # Listen für Nachweise
    L_Md = [M_yd, M_zd]
    lamb_rel = min(L_lamb_rel)
    L_pot = [[1, 2], [2, 1]]
    k_crit_index = L_w.index(max(L_w))

    # Schleife für Nachweise
    for n, M in enumerate(L_Md):
        # EC5 Abs. 6.2.4: Biegung m/o Druck ohne Knicken ohne Kippen
        if k_crit == 1 and lamb_rel < 0.3:
            eta = (sigma_cd/f_c0d)**2 + \
                L_km[n][0]*L_sigma_md[0]/f_myd + L_km[n][1]*L_sigma_md[1]/f_mzd
            L_eta.append(round(eta, 2))
            nw = ('Sp')

        # EC5 Abs. 6.3.2: Biegung m/o Druck mit Knicken ohne Kippen
        elif k_crit == 1 and lamb_rel > 0.3:
            eta = (sigma_cd)/(f_c0d*L_kc[n]) + L_km[n][0] * \
                L_sigma_md[0]/f_myd + L_km[n][1]*L_sigma_md[1]/f_mzd
            L_eta.append(round(eta, 2))
            nw = ('Kn')

        # EC5 Abs. 6.3.3: Biegung m/o Druck mit Knicken und Kippen
        elif k_crit < 1 and lamb_rel < 0.3:
            eta = (sigma_cd)/(f_c0d*L_kc[n]) + (L_sigma_md[0]/(f_myd*L_k_crit[0])
                                                )**L_pot[n][0] + (L_sigma_md[1]/(f_mzd*L_k_crit[1]))**L_pot[n][1]
            L_eta.append(round(eta, 2))
            nw = ('Kn/Ki')

        else:
            eta = 0
            L_eta.append(round(eta, 2))
            nw = ('N/A')

    return e_0, L_lamb, L_lamb_rel, L_ky, L_kc, k_crit, sigma_cd, L_sigma_md, L_eta, nw

# Stützenbemessungsprogramm: Excel-Iteration - Ersatzstabverfahren
def ec5_esv_iteration():

    # Excel
    wb = xw.Book.caller()  # Arbeitsmappe
    ws = wb.sheets.active  # Arbeitsblatt

    startrow = 10  # Startzeile
    # Dynamische Anzahl an Reihen
    rownum = ws.range('G10').current_region.last_cell.row
    ws.range('X10:AH100').value = ""  # Alte Daten Löschen

    # Einlesen der Listen
    L_pos = ws.range((startrow, 7), (rownum, 7)).value
    L_ges = ws.range((startrow, 8), (rownum, 8)).value
    L_system = ws.range((startrow, 9), (rownum, 9)).value
    L_L = ws.range((startrow, 10), (rownum, 10)).value
    L_b = ws.range((startrow, 11), (rownum, 11)).value
    L_h = ws.range((startrow, 12), (rownum, 12)).value
    L_N_ed = ws.range((startrow, 13), (rownum, 13)).value
    L_M_yd = ws.range((startrow, 14), (rownum, 14)).value
    L_M_zd = ws.range((startrow, 15), (rownum, 15)).value
    L_h_art = ws.range((startrow, 16), (rownum, 15)).value
    L_f_c0k = ws.range((startrow, 17), (rownum, 17)).value
    L_f_myk = ws.range((startrow, 18), (rownum, 18)).value
    L_E0mean = ws.range((startrow, 19), (rownum, 19)).value
    L_E_05 = ws.range((startrow, 20), (rownum, 20)).value
    L_G_05 = ws.range((startrow, 21), (rownum, 21)).value
    L_k_mod = ws.range((startrow, 22), (rownum, 22)).value
    L_gamma = ws.range((startrow, 23), (rownum, 23)).value

    # Schleife: Iteration durch Stützenpositionen
    for i in range(rownum - startrow + 1):

        # Definition der Variablen des i-ten Elements der Listen
        L = L_L[i]
        b = L_b[i]
        h = L_h[i]
        güte = L_h_art[i]
        Lagerung = L_system[i]
        f_c0k = L_f_c0k[i]*1000
        f_myk = L_f_myk[i]*1000
        f_mzk = f_myk
        E0_mean = L_E0mean[i]*1000
        E_05 = L_E_05[i]*1000
        G_05 = L_G_05[i]*1000
        k_mod = L_k_mod[i]
        gamma = L_gamma[i]
        N_ed = L_N_ed[i]
        M_yd = L_M_yd[i]
        M_zd = L_M_zd[i]

        # Berechnung
        # Kennwerte
        theta = 0
        shi = k_mod/gamma
        f_c0d = f_c0k*shi
        f_myd = f_myk*shi
        E = E_05

        # Stützenbemessung nach Theorie I. Ordnung (Funktion)
        e_0, L_lamb, L_lamb_rel, L_ky, L_kc, k_crit, sigma_cd, L_sigma_md, L_eta, nw = ec5_63_esv(
            Lagerung, güte, L, b, h, N_ed, M_yd, M_zd, theta, f_c0k, f_myk, f_mzk, E, G_05, k_mod, gamma)

        # Auswertung der Ergebnisse
        lamb = max(L_lamb)
        index = L_lamb.index(lamb)
        lamb_rel = max(L_lamb_rel)
        kc = min(L_kc)
        sigma_myd = L_sigma_md[0]
        sigma_mzd = L_sigma_md[1]
        eta = max(L_eta)

        #Schreiben in Excel
        ws.range('Y10').offset(i, 0).value = round(lamb, 2)
        ws.range('Z10').offset(i, 0).value = round(lamb_rel, 2)
        ws.range('AA10').offset(i, 0).value = round(kc, 2)
        ws.range('AB10').offset(i, 0).value = round(k_crit, 2)
        ws.range('AC10').offset(i, 0).value = nw
        ws.range('AD10').offset(i, 0).value = sigma_cd
        ws.range('AE10').offset(i, 0).value = round(sigma_myd, 2)
        ws.range('AF10').offset(i, 0).value = round(sigma_mzd, 2)
        ws.range('AG10').offset(i, 0).value = round(eta, 2)

# Stützenbemessungsprogramm: Funktion - Theorie II. Ordnung
def ec5_63_th2(Lagerung, güte, L, b, h, N_ed, M_yd, M_zd, theta, f_c0k, f_myk, f_mzk, E, k_mod, gamma, no_iter):

    # Listen
    L_lamb = []
    L_eta = []

    # Bemessungswerte der Festigkeit
    xi = k_mod/gamma
    f_c0d = f_c0k*xi  # $N/mm^2$ - Bemessungswert der Druckfestigkeit
    f_myd = f_myk*xi  # $N/mm^2$ - Bemessungswert der Biegefestigkeit
    f_mzd = f_mzk*xi  # $N/mm^2$ - Bemessungswert der Biegefestigkeit

    # Beiwert k_m
    if M_yd == 0 or M_zd == 0:
        L_km = [[1, 1], [1, 1]]
    else:
        L_km = [[1, 0.7], [0.7, 1]]

    # if-Abfrage zur Bestimmung von beta
    if Lagerung == 'Pendelstütze':
        beta = 1
    elif Lagerung == 'Kragstütze':
        beta = 2
    elif Lagerung == 'Eingespannt (u/o)':
        beta = 0.7
    else:
        beta = 0.5

    # Ausmitte nach Theorie I. Ordnung
    l_ef = L*beta
    e_0 = round(theta*l_ef, 8)
    M_0 = e_0 * N_ed

    # Listen
    L_bh = [[b, h], [h, b], [b, h], [h, b]]
    L_e = [[e_0*1000], [0], [0], [e_0*1000]]
    L_M = [[M_yd+M_0], [M_zd], [M_yd], [M_zd+M_0]]
    L_e_total = [[e_0*1000], [0], [0], [e_0*1000]]
    L_M_total = [[M_yd+M_0], [M_zd], [M_yd], [M_zd+M_0]]
    L_sigma_mIId = []

    # Schleife 1: Durchlaufen von [M_yIId_imp, M_zIId, M_yIId, M_zIId_imp ]
    for n in range(4):

        # Querschnittsparameter
        A = b*h  # $m^2$
        I = (L_bh[n][0]*L_bh[n][1]**3)/12  # $m^4$ - FTM
        w = (L_bh[n][0]*L_bh[n][1]**2)/6  # $m^3$ - Widerstandsmoment
        i = L_bh[n][1]/sqrt(12)  # $m$ - polares Trägheitsmoment

        # Schleife 2: Schnittgrößenermittlung nach Theorie II. Ordnung
        for i in range(no_iter):

            # Ermittlung der Werte
            #e_i = (L_M[n][i]*l_ef**2)/(E*I*pi**2)
            e_i = (40*L_M[n][i]*L**2)/(E*I*384)
            M_i = e_i * N_ed

            # Anhängen der Werte in Listen
            L_e[n].append(round(e_i*1000, 1))
            L_M[n].append(M_i)
            e_total = sum(L_e[n])
            M_total = sum(L_M[n])
            L_e_total[n].append(round(e_total, 2))
            L_M_total[n].append(round(M_total, 2))

        # Spannungen
        sigma_mIId = L_M_total[n][-1]/w  # $kN/m^2$
        L_sigma_mIId.append(round(sigma_mIId, 2))

    # Ergebnisse der Schnittgrößenermittlung nach Theorie II. Ordnung
    # Verformungen und Momente
    L_e = [[L_e[0], L_e[1]], [L_e[2], L_e[3]]]
    L_M = [[L_M[0], L_M[1]], [L_M[2], L_M[3]]]
    L_e_total = [[L_e_total[0], L_e_total[1]], [L_e_total[2], L_e_total[3]]]
    L_M_total = [[L_M_total[0], L_M_total[1]], [L_M_total[2], L_M_total[3]]]

    # Spannungen
    sigma_cd = N_ed/A  # $kN/m^2$ - Druckspannung
    L_sigma_mIId = [[L_sigma_mIId[0], L_sigma_mIId[1]],
                    [L_sigma_mIId[2], L_sigma_mIId[3]]]

    # Schleife für y- und z-Achse
    for i in range(2):

        # Nachweis
        eta = (sigma_cd/f_c0d)**2 + \
            L_km[i][0]*L_sigma_mIId[i][0]/f_myd + \
            L_km[i][1]*L_sigma_mIId[i][1]/f_mzd
        L_eta.append(round(eta, 6))

    # maximale Ausnutzung
    eta_max = max(L_eta)
    index = L_eta.index(eta_max)

    return L_e[index], L_M[index], L_e_total[index], L_M_total[index], sigma_cd, L_sigma_mIId[index], eta_max, index


# Stützenbemessungsprogramm: Excel-Iteration - Theorie II. Ordnung
def ec5_th2_iteration():

    # Excel
    wb = xw.Book.caller()
    ws = wb.sheets.active
    
    # Dynamisches Einlesen der Zeilenanzahl
    startrow = 10
    rownum = ws.range('G10').current_region.last_cell.row
    colnum = ws.range((startrow, 1)).current_region.last_cell.column

    # Alte Daten löschen
    ws.range('AI10:AQ100').value = ""

    # Einlesen der Listen
    L_pos = ws.range((startrow, 7), (rownum, 7)).value
    L_ges = ws.range((startrow, 8), (rownum, 8)).value
    L_system = ws.range((startrow, 9), (rownum, 9)).value
    L_L = ws.range((startrow, 10), (rownum, 10)).value
    L_b = ws.range((startrow, 11), (rownum, 11)).value
    L_h = ws.range((startrow, 12), (rownum, 12)).value
    L_N_ed = ws.range((startrow, 13), (rownum, 13)).value
    L_M_yd = ws.range((startrow, 14), (rownum, 14)).value
    L_M_zd = ws.range((startrow, 15), (rownum, 15)).value
    L_h_art = ws.range((startrow, 16), (rownum, 15)).value
    L_f_c0k = ws.range((startrow, 17), (rownum, 17)).value
    L_f_myk = ws.range((startrow, 18), (rownum, 18)).value
    L_E0mean = ws.range((startrow, 19), (rownum, 19)).value
    L_E_05 = ws.range((startrow, 20), (rownum, 20)).value
    L_G_05 = ws.range((startrow, 21), (rownum, 21)).value
    L_k_mod = ws.range((startrow, 22), (rownum, 22)).value
    L_gamma = ws.range((startrow, 23), (rownum, 23)).value

    e_0 = ws.range('D24').value

    for j in range(rownum - startrow + 1):

        # Definition der Variablen des j-ten Elements der Listen
        L = L_L[j]
        b = L_b[j]
        h = L_h[j]
        f_c0k = L_f_c0k[j]*1000
        f_myk = L_f_myk[j]*1000
        f_mzk = f_myk
        E0_mean = L_E0mean[j]*1000
        E_05 = L_E_05[j]*1000
        k_mod = L_k_mod[j]
        gamma = L_gamma[j]
        N_ed = L_N_ed[j]
        M_yd = L_M_yd[j]
        M_zd = L_M_zd[j]

        güte = L_h_art[j]
        Lagerung = L_system[j]
        no_iter = 5
        E = E0_mean/1.3

        # Kennwerte
        shi = k_mod/gamma
        f_c0d = f_c0k*shi
        f_myd = f_myk*shi

        A = b*h
        I = b*h**3/12
        w = (b*h**2)/6
        i = (I/A)**0.5

        # Stützenbemessung nach Theorie I. Ordnung (Funktion)
        L_e, L_M, L_e_total, L_M_total, sigma_cd, L_sigma_mIId, eta, index = ec5_63_th2(
            Lagerung, güte, L, b, h, N_ed, M_yd, M_zd, e_0, f_c0k, f_myk, f_mzk, E, k_mod, gamma, no_iter)

        #Schreiben in Excel
        ws.range('AI10').offset(j, 0).value = L_e_total[0][0]
        ws.range('AJ10').offset(j, 0).value = L_M_total[0][0]
        ws.range('AK10').offset(j, 0).value = L_e_total[0][-1]
        ws.range('AL10').offset(j, 0).value = L_M_total[0][-1]
        ws.range('AM10').offset(j, 0).value = L_e[1][0]
        ws.range('AN10').offset(j, 0).value = L_M[1][0]
        ws.range('AO10').offset(j, 0).value = L_e_total[1][-1]
        ws.range('AP10').offset(j, 0).value = L_M_total[1][-1]
        ws.range('AQ10').offset(j, 0).value = eta

# Stützenbemessungsprogramm: Optimierungsalgorithmus
def ec5_optimierung():

    # Excel
    wb = xw.Book('Stützenbemessung_ec5_63.xlsm')
    ws = wb.sheets['Stützenbemessung']
    # Startreihe
    startrow = 10
    rownum = ws.range('G10').current_region.last_cell.row
    colnum = ws.range((startrow, 1)).current_region.last_cell.column
    n = rownum-startrow

    # Alte Daten löschen
    for i in range(n-1):
        if ws.range('BH10').offset(i, 0).value == 'NW n. erbracht':
            ws.range((startrow+i, 45), (startrow+i, 61)).value = ''
        else:
            print('')

    # Status
    ws.range('BH8').value = 'läuft..'

    # Optimierung
    # Ausnutzung
    eta_target_top = ws.range('D37').value
    eta_target_bot = ws.range('D38').value

    # Querschnitt
    para1_target_top = ws.range('D32').value
    para1_target_bot = ws.range('D33').value
    para1_steps = ws.range('D34').value
    n_steps = (para1_target_top-para1_target_bot)/para1_steps

    # Güte
    para2_target_top = ws.range('D28').value
    para2_target_bot = ws.range('D29').value

    L_güte = ['GL 24h', 'GL 24c', 'GL 28h', 'GL 28c', 'GL 32h', 'GL 32c']
    index_top = L_güte.index(para2_target_top)
    index_bot = L_güte.index(para2_target_bot)

    # Einlesen der Listen
    L_pos = ws.range((startrow, 7), (rownum, 7)).value
    L_ges = ws.range((startrow, 8), (rownum, 8)).value
    L_system = ws.range((startrow, 9), (rownum, 9)).value
    L_L = ws.range((startrow, 10), (rownum, 10)).value
    L_b = ws.range((startrow, 11), (rownum, 11)).value
    L_h = ws.range((startrow, 12), (rownum, 12)).value
    L_N_ed = ws.range((startrow, 13), (rownum, 13)).value
    L_M_yd = ws.range((startrow, 14), (rownum, 14)).value
    L_M_zd = ws.range((startrow, 15), (rownum, 15)).value
    L_h_art = ws.range((startrow, 16), (rownum, 16)).value
    L_f_c0k = ws.range((startrow, 17), (rownum, 17)).value
    L_f_myk = ws.range((startrow, 18), (rownum, 18)).value
    L_E0mean = ws.range((startrow, 19), (rownum, 19)).value
    L_E_05 = ws.range((startrow, 20), (rownum, 20)).value
    L_G_05 = ws.range((startrow, 21), (rownum, 21)).value
    L_k_mod = ws.range((startrow, 22), (rownum, 22)).value
    L_gamma = ws.range((startrow, 23), (rownum, 23)).value

    a1 = 'NW erfüllt'
    a2 = 'NW n. erbracht'
    a3 = 'Grenze'
    L_n_iter = []
    L_n_erf = []
    L_n_nerf = []

    # Starten der Schleife
    # j für Reihen
    for j in range(rownum - startrow + 1):
        # for j in range(30):
        ws.range('BH9').value = str(j) + '/' + str(n)
        güte = L_h_art[j]

        # i für Spalten
        for k, i in enumerate(range(5)):

            # Rechnet nur die leeren Positionen
            if ws.range('BH10').offset(j, 0).value == a1:
                break

            elif ws.range('BG10').offset(j, 0).value == None or ws.range('BH10').offset(j, 0).value == a2:

                # Definition der Variablen des i-ten Elements der Listen
                L = L_L[j]
                Lagerung = L_system[j]
                k_mod = L_k_mod[j]
                gamma = L_gamma[j]
                N_ed = L_N_ed[j]
                M_yd = L_M_yd[j]
                M_zd = L_M_zd[j]

                b = ws.range('K10').offset(j, 0).value
                h = ws.range('L10').offset(j, 0).value

                güte = ws.range('P10').offset(j, 0).value
                index = L_güte.index(güte)
                f_c0k = ws.range('Q10').offset(j, 0).value*1000
                f_myk = ws.range('R10').offset(j, 0).value*1000
                f_mzk = f_myk
                E0_mean = ws.range('S10').offset(j, 0).value*1000
                E_05 = ws.range('T10').offset(j, 0).value*1000
                G_05 = ws.range('U10').offset(j, 0).value*1000

                # Theorie I. Ordnung
                theta = 0
                E = E_05

                # Funktion
                e_0, L_lamb, L_lamb_rel, L_ky, L_kc, k_crit, sigma_cd, L_sigma_md, L_eta, nw = ec5_63_esv(
                    Lagerung, güte, L, b, h, N_ed, M_yd, M_zd, theta, f_c0k, f_myk, f_mzk, E, G_05, k_mod, gamma)

                # Auswertung der Ergebnisse
                lamb = max(L_lamb)
                lamb_rel = max(L_lamb_rel)
                kc = min(L_kc)
                sigma_myd = L_sigma_md[0]
                sigma_mzd = L_sigma_md[1]
                eta_esv = max(L_eta)

                # Theorie II. Ordnung
                theta = 0.0025
                E = E0_mean/1.3
                no_iter = 5

                # Funktion
                L_e, L_M, L_e_total, L_M_total, sigma_cd, L_sigma_mIId, eta_th2, ii = ec5_63_th2(
                    Lagerung, güte, L, b, h, N_ed, M_yd, M_zd, theta, f_c0k, f_myk, f_mzk, E, k_mod, gamma, no_iter)

                # Auswertung der Ergebnisse
                L_e_y = L_e[0]
                L_M_y = L_M[0]
                L_e_total_y = L_e_total[0]
                L_M_total_y = L_M_total[0]

                L_e_z = L_e[1]
                L_M_z = L_M[1]
                L_e_total_z = L_e_total[1]
                L_M_total_z = L_e_total[1]

                # maximale Ausnutzung
                eta = max(eta_esv, eta_th2)
                ws.range('AS10').offset(j, i).value = güte
                ws.range('AX10').offset(j, i).value = round(b, 2)
                ws.range('BC10').offset(j, i).value = round(eta, 2)

                # FALL 1: Iteration fertig / Ausnutzung bereits erreicht
                if i == 4 and (eta > eta_target_top or eta < eta_target_bot):

                    # Status
                    ws.range('BH10').offset(j, 0).value = a2

                    # ESV
                    ws.range('Y10').offset(j, 0).value = round(lamb, 2)
                    ws.range('Z10').offset(j, 0).value = round(lamb_rel, 2)
                    ws.range('AA10').offset(j, 0).value = round(kc, 2)
                    ws.range('AB10').offset(j, 0).value = round(k_crit, 2)
                    ws.range('AC10').offset(j, 0).value = nw
                    ws.range('AD10').offset(j, 0).value = sigma_cd
                    ws.range('AE10').offset(j, 0).value = round(sigma_myd, 2)
                    ws.range('AF10').offset(j, 0).value = round(sigma_mzd, 2)
                    ws.range('AG10').offset(j, 0).value = round(eta_esv, 2)

                    # TH2
                    ws.range('AI10').offset(j, 0).value = L_e[0][0]
                    ws.range('AJ10').offset(j, 0).value = L_M[0][0]
                    ws.range('AK10').offset(j, 0).value = L_e_total[0][-1]
                    ws.range('AL10').offset(j, 0).value = L_M_total[0][-1]

                    ws.range('AM10').offset(j, 0).value = L_e[1][0]
                    ws.range('AN10').offset(j, 0).value = L_M[1][0]
                    ws.range('AO10').offset(j, 0).value = L_e_total[1][-1]
                    ws.range('AP10').offset(j, 0).value = L_M_total[1][-1]
                    ws.range('AQ10').offset(j, 0).value = eta_th2

                # FALL 2: Ausnutzung erreicht
                elif eta <= eta_target_top and eta >= eta_target_bot:

                    # ESV
                    ws.range('Y10').offset(j, 0).value = round(lamb, 2)
                    ws.range('Z10').offset(j, 0).value = round(lamb_rel, 2)
                    ws.range('AA10').offset(j, 0).value = round(kc, 2)
                    ws.range('AB10').offset(j, 0).value = round(k_crit, 2)
                    ws.range('AC10').offset(j, 0).value = nw
                    ws.range('AD10').offset(j, 0).value = sigma_cd
                    ws.range('AE10').offset(j, 0).value = round(sigma_myd, 2)
                    ws.range('AF10').offset(j, 0).value = round(sigma_mzd, 2)
                    ws.range('AG10').offset(j, 0).value = round(eta_esv, 2)

                    # TH2
                    ws.range('AI10').offset(j, 0).value = L_e[0][0]
                    ws.range('AJ10').offset(j, 0).value = L_M[0][0]
                    ws.range('AK10').offset(j, 0).value = L_e_total[0][-1]
                    ws.range('AL10').offset(j, 0).value = L_M_total[0][-1]

                    ws.range('AM10').offset(j, 0).value = L_e[1][0]
                    ws.range('AN10').offset(j, 0).value = L_M[1][0]
                    ws.range('AO10').offset(j, 0).value = L_e_total[1][-1]
                    ws.range('AP10').offset(j, 0).value = L_M_total[1][-1]
                    ws.range('AQ10').offset(j, 0).value = eta_th2

                    # PARAMETER
                    ws.range('K10').offset(j, 0).value = round(b, 2)
                    ws.range('L10').offset(j, 0).value = round(h, 2)
                    ws.range('P10').offset(j, 0).value = güte

                    ws.range('AS10').offset(j, i).value = güte
                    ws.range('AX10').offset(j, i).value = round(b, 2)
                    ws.range('BC10').offset(j, i).value = round(eta, 2)

                    # ITERATION
                    ws.range('BH10').offset(j, 0).value = a1
                    break

                # FALL 3: Ausnutzung über 100%
                elif eta > 1:
                    # Güte: Grenze bereits erreicht
                    if güte == para2_target_top:
                        # Querschnitt raufgesetzt
                        b = min(b+2*para1_steps, para1_target_top)
                        h = b
                        if b == para1_target_top and i < 4:
                            ws.range('BH10').offset(j, 0).value = a3
                            break

                    else:
                        # Querschnitt bleibt / Güte wird raufgesetzt
                        b = b
                        h = h
                        güte = L_güte[min(index+1, index_top)]

                # FALL 4: Ausnutzung über eta_top
                elif eta > eta_target_top:
                    # Güte: Grenze bereits erreicht
                    if güte == para2_target_top:
                        # Querschnitt raufgesetzt
                        b = min(b+para1_steps, para1_target_top)
                        h = b

                        if b == para1_target_top and i < 4:
                            ws.range('BH10').offset(j, 0).value = a3
                            break

                    else:
                        # Querschnitt bleibt / Güte wird raufgesetzt
                        b = b
                        h = h
                        güte = L_güte[min(index+1, index_top)]

                # FALL 5: Ausnutzung unter eta_bot
                elif eta < eta_target_bot:
                    # Güte: Grenze bereits erreicht
                    if güte == para2_target_bot:
                        b = max(b-para1_steps, para1_target_bot)
                        h = b

                        if b == para1_target_bot and i < 4:
                            ws.range('BH10').offset(j, 0).value = a3
                            break
                    else:
                        b = b
                        h = h
                        güte = L_güte[max(index-1, index_bot)]

                # ITERATION
                ws.range('P10').offset(j, 0).value = güte
                ws.range('K10').offset(j, 0).value = round(b, 2)
                ws.range('L10').offset(j, 0).value = round(h, 2)

            else:
                continue

    # Status
    ws.range('BH8').value = 'abgeschlossen'

# Stützenbemessungsprogramm: Importieren der Daten aus dem Lastabtrag
def ec5_import_data():

    wb = xw.Book('Stützenbemessung_ec5_63.xlsm')
    ws_source = wb.sheets['Lastabtrag']
    ws_target = wb.sheets.active
    geschoss = ws_target.range('D13').value

    # Default
    startrow = 10
    güte = 'GL 24h'
    system = 'Pendelstütze'
    kmod = 0.6
    gamma = 1.3

    rownum = ws_source.range('B10').current_region.last_cell.row
    geschosse = ws_source.range((startrow, 3), (rownum, 3)).value
    start_geschoss = geschosse.index(geschoss)
    index = []

    for i, f in enumerate(geschosse):
        if f in geschoss:
            index.append(i)
    rownum_geschoss = index[-1]-start_geschoss

    ws_target.range('G10:P200').value = ''
    ws_target.range('V10:AQ200').value = ''
    ws_target.range('AS10:BI200').value = ''

    ws_target.range((startrow, 7), (rownum_geschoss+startrow, 8)).value = ws_source.range(
        (startrow+start_geschoss, 2), (startrow+start_geschoss+rownum_geschoss, 3)).value
    ws_target.range((startrow, 10), (rownum_geschoss+startrow, 12)).value = ws_source.range(
        (startrow+start_geschoss, 6), (startrow+start_geschoss+rownum_geschoss, 8)).value
    ws_target.range((startrow, 13), (rownum_geschoss+startrow, 15)).value = ws_source.range(
        (startrow+start_geschoss, 9), (startrow+start_geschoss+rownum_geschoss, 11)).value
    ws_target.range((startrow, 9), (rownum_geschoss +
                    startrow, 9)).value = system
    ws_target.range(
        (startrow, 16), (rownum_geschoss+startrow, 16)).value = güte
    ws_target.range(
        (startrow, 22), (rownum_geschoss+startrow, 22)).value = kmod
    ws_target.range((startrow, 23), (rownum_geschoss +
                    startrow, 23)).value = gamma

# Stützenbemessungsprogramm: Aufrunden der Einwirkungen
def ec5_group_loads():

    # Excel
    wb = xw.Book('Stützenbemessung_ec5_63.xlsm')
    ws = wb.sheets.active
    # Startreihe
    startrow = 10
    rownum = ws.range('G10').current_region.last_cell.row
    colnum = ws.range((startrow, 1)).current_region.last_cell.column
    n = rownum-startrow

    # Liste für Einwirkungen
    L_N_ed = ws.range((startrow, 13), (rownum, 13)).value
    L_M_yd = ws.range((startrow, 14), (rownum, 14)).value
    L_M_zd = ws.range((startrow, 15), (rownum, 15)).value

    # Rundungsparameter
    round_ned = ws.range('D18').value
    round_med = ws.range('D19').value

    # Listen
    L_Ned_round = []
    L_Myd_round = []
    L_Mzd_round = []

    # Funktion
    def roundup(numb, value):
        return int(ceil(abs(numb) / value) * value)

    for i in range(n-1):

        N_ed_round = roundup(L_N_ed[i], round_ned)
        Myd_round = 0 if L_M_yd[i] == 0 else roundup(L_M_yd[i], round_med)
        Mzd_round = 0 if L_M_zd[i] == 0 else roundup(L_M_zd[i], round_med)

        ws.range('M10').offset(i, 0).value = N_ed_round
        ws.range('N10').offset(i, 0).value = Myd_round
        ws.range('O10').offset(i, 0).value = Mzd_round

# Stützenbemessungsprogramm: Dokumentation - Grafische Übersicht mit AutoCAD
def docu_cad():
    try:

        # Excel
        wb = xw.Book('Stützenbemessung_ec5_63.xlsm')
        ws_1 = wb.sheets['Lastabtrag']
        ws_2 = wb.sheets['Stützenbemessung']

        startrow = 10
        rownum = ws_2.range('G10').current_region.last_cell.row

        # Koordinaten ws_2
        # Stützen
        x_coord = ws_1.range((10, 13), (rownum, 13)).value
        y_coord = ws_1.range((10, 14), (rownum, 14)).value

        # Gebäudeaußenkante
        xx_coord = ws_1.range((10, 17), (22, 17)).value
        yy_coord = ws_1.range((10, 18), (22, 18)).value

        # Stützenkennwerte ws_1
        pos = ws_2.range((startrow, 7), (rownum, 7)).value
        geschoss = ws_2.range((startrow, 8), (rownum, 8)).value
        material = ws_2.range((startrow, 16), (rownum, 16)).value

        L = ws_2.range((startrow, 10), (rownum, 10)).value
        b = ws_2.range((startrow, 11), (rownum, 11)).value
        h = ws_2.range((startrow, 12), (rownum, 12)).value

        Ned = ws_2.range((startrow, 13), (rownum, 13)).value
        Myd = ws_2.range((startrow, 14), (rownum, 14)).value
        Mzd = ws_2.range((startrow, 15), (rownum, 15)).value

        # Auswertung
        # Güte
        L_güte = ['GL 24h', 'GL 24c', 'GL 28h', 'GL 28c', 'GL 32h', 'GL 32c']
        L_color_güte = [1, 1, 4, 4, 3, 3]

        # Querschnitt
        L_qs = []
        b_oben = ws_2.range('D32').value
        b_unten = ws_2.range('D33').value
        b_steps = ws_2.range('D34').value
        n_steps = (b_oben-b_unten)/b_steps
        L_color_qs = []

        for i in range(int(n_steps)):
            L_qs.append(round(b_unten+b_steps*i, 2))
            L_color_qs.append(i)

        L_qs[-1] = b_oben

        # AutoCAD
        # Öffnen einer neuen Datei
        acad = Autocad(create_if_not_exists=True)

        # Zeichnen der Decke
        for i in range(0, 12):
            p1 = APoint(xx_coord[i], yy_coord[i])  # Punkt 1
            p2 = APoint(xx_coord[i+1], yy_coord[i+1])  # Punkt 2
            line1 = acad.model.AddLine(p1, p2)  # Linie zwischen Punkt 1 und 2

        # Stützen
        for i in range(len(x_coord)):

            # Index für Güte und Querschnitt
            index_güte = L_güte.index(material[i])
            index_qs = L_qs.index(round(b[i], 2))

            # Punktkoordinaten
            pi = APoint(x_coord[i], y_coord[i])
            p1 = APoint(x_coord[i]+b[i]*0.5, y_coord[i]+h[i]*0.5)
            p2 = APoint(x_coord[i]+b[i]*0.5, y_coord[i]-h[i]*0.5)
            p3 = APoint(x_coord[i]-b[i]*0.5, y_coord[i]-h[i]*0.5)
            p4 = APoint(x_coord[i]-b[i]*0.5, y_coord[i]+h[i]*0.5)

            # Zeichnen der Stütze
            line1 = acad.model.AddLine(p1, p2)
            line2 = acad.model.AddLine(p2, p3)
            line3 = acad.model.AddLine(p3, p4)
            line4 = acad.model.AddLine(p4, p1)

            # Anpassung der Farbe gemäß Index
            line1.Color = L_color_qs[index_qs]
            line2.Color = L_color_qs[index_qs]
            line3.Color = L_color_qs[index_qs]
            line4.Color = L_color_qs[index_qs]

            # Textelemente
            text = acad.model.AddText(pos[i] + ' - ' + geschoss[i], pi, 0.2)
            text1 = acad.model.AddText(
                material[i], APoint(x_coord[i], y_coord[i]-0.3), 0.2)
            text2 = acad.model.AddText('L/B/H = '+str(round(L[i], 2))+'/'+str(
                b[i]) + '/' + str(h[i])+' [m]', APoint(x_coord[i], y_coord[i]-0.6), 0.2)
            text3 = acad.model.AddText('Ned/Myd/Mzd = '+str(round(Ned[i], 2))+'/'+str(
                Myd[i]) + '/' + str(Mzd[i]), APoint(x_coord[i], y_coord[i]-0.9), 0.2)

            # Anpassung der Farbe gemäß Index
            text1.Color = L_color_güte[index_güte]

    except:
        print('')

# Stützenbemessungsprogramm: Dokumentation - Schriftlicher Einzelnachweis als PDF
def docu_pdf():

    # Verknüpfung mit Excel
    wb = xw.Book('Stützenbemessung_ec5_63.xlsm')
    ws_1 = wb.sheets['Stützenbemessung']
    ws_2 = wb.sheets['Theorie II. Ordnung']

    # Dynamisches Einlesen der Zeilenanzahl
    startrow = 10
    rownum = ws_1.range('G10').current_region.last_cell.row
    colnum = ws_1.range((startrow, 1)).current_region.last_cell.column

    # Einlesen der Listen
    L_pos = ws_1.range((startrow, 7), (rownum, 7)).value
    L_ges = ws_1.range((startrow, 8), (rownum, 8)).value
    L_system = ws_1.range((startrow, 9), (rownum, 9)).value
    L_L = ws_1.range((startrow, 10), (rownum, 10)).value
    L_b = ws_1.range((startrow, 11), (rownum, 11)).value
    L_h = ws_1.range((startrow, 12), (rownum, 12)).value
    L_N_ed = ws_1.range((startrow, 13), (rownum, 13)).value
    L_M_yd = ws_1.range((startrow, 14), (rownum, 14)).value
    L_M_zd = ws_1.range((startrow, 15), (rownum, 15)).value
    L_h_art = ws_1.range((startrow, 16), (rownum, 15)).value
    L_f_c0k = ws_1.range((startrow, 17), (rownum, 17)).value
    L_f_myk = ws_1.range((startrow, 18), (rownum, 18)).value
    L_E0mean = ws_1.range((startrow, 19), (rownum, 19)).value
    L_E_05 = ws_1.range((startrow, 20), (rownum, 20)).value
    L_G_05 = ws_1.range((startrow, 21), (rownum, 21)).value
    L_k_mod = ws_1.range((startrow, 22), (rownum, 22)).value
    L_gamma = ws_1.range((startrow, 23), (rownum, 23)).value

    # Schleife durch Stützenpositionen
    for j in range(rownum - startrow + 1):

        # Ermittlung der gekennzeichneten Positionen
        if ws_1.range("BJ10").offset(j, 0).value == "x":

            # Einlesen der Kennwerte gemäß Index
            L = L_L[j]
            b = L_b[j]
            h = L_h[j]
            f_c0k = L_f_c0k[j]*1000
            f_myk = L_f_myk[j]*1000
            f_mzk = f_myk
            E0_mean = L_E0mean[j]*1000
            E_05 = L_E_05[j]*1000
            k_mod = L_k_mod[j]
            gamma = L_gamma[j]
            N_ed = L_N_ed[j]
            M_yd = L_M_yd[j]
            M_zd = L_M_zd[j]
            güte = L_h_art[j]
            Lagerung = L_system[j]
            theta = ws_1.range("D24").value
            no_iter = 9
            E = E0_mean/1.3

            # Auswertung der Kennwerte
            shi = k_mod/gamma
            f_c0d = f_c0k*shi
            f_myd = f_myk*shi

            A = b*h
            I = b*h**3/12
            w = (b*h**2)/6
            i = (I/A)**0.5

            # Bemessung mit Funktion
            L_e, L_M, L_e_total, L_M_total, sigma_cd, L_sigma_mIId, eta, index = ec5_63_th2(
                Lagerung, güte, L, b, h, N_ed, M_yd, M_zd, theta, f_c0k, f_myk, f_mzk, E, k_mod, gamma, no_iter)

            #Schreiben in bemessungsvorlage
            # Geometrie
            ws_2.range('I4').value = L_pos[j]
            ws_2.range('T49').value = index

            # Geometrie
            start = ws_2.range('E14')
            start.offset(0, 0).value = L
            start.offset(0, 13).value = b
            start.offset(0, 26).value = h

            # Bemessungswerte der Einwirkung
            start = ws_2.range('E17')
            start.offset(0, 0).value = N_ed
            start.offset(0, 13).value = M_yd
            start.offset(0, 26).value = M_zd

            # Charakteristische Festigkeitswerte
            start = ws_2.range('F21')
            start.offset(0, 0).value = f_c0k
            start.offset(0, 16).value = f_myk
            start.offset(1, 0).value = E0_mean

            # Widerstandsbeiwerte
            start = ws_2.range('F25')
            start.offset(0, 0).value = k_mod
            start.offset(0, 16).value = gamma

            # Charakteristische Festigkeitswerte
            start = ws_2.range('F34')
            start.offset(0, 0).value = f_c0d
            start.offset(0, 16).value = f_myd
            start.offset(1, 0).value = E0_mean/1.3

            # Schnittgrößenermittlung nach Theorie II. Ordnung
            start = ws_2.range('E61')
            start.offset(0, 0).options(transpose=True).value = L_e[0]
            start.offset(0, 6).options(transpose=True).value = L_M[0]
            start.offset(0, 12).options(transpose=True).value = L_e_total[0]
            start.offset(0, 18).options(transpose=True).value = L_M_total[0]

            start = ws_2.range('E80')
            start.offset(0, 0).options(transpose=True).value = L_e[1]
            start.offset(0, 6).options(transpose=True).value = L_M[1]
            start.offset(0, 12).options(transpose=True).value = L_e_total[1]
            start.offset(0, 18).options(transpose=True).value = L_M_total[1]

            # PDF export
            pfad = ws_1.range("D41").value
            pdf_file_name = pfad + "Stützenbemessung_ThII_" + L_pos[j]
            ws_2.to_pdf(pdf_file_name)
