"""Cost utility functions for air separation plant equipment and economics."""
import numpy as np


def pump_cost(W):
    effi = 0.8
    return 705.48 * W ** 0.71 * (1 + 0.2 / (1 - effi))


def ASU_cost(Size):
    RefCost = 141  # million 2014
    RefSize = 52  # kg O2/sec
    Scaling = 0.5
    Installation = 1
    return RefCost * Installation * (Size / RefSize) ** Scaling


def Dist_C(L, d, P, FM, Nact, fq):
    FBM = 1
    CVe = 1780 * (L ** 0.87) * (d ** 1.23) * (
        2.86 + 1.694 * FM * (10.011 - 7.408 * np.log(P) + 1.395 * (np.log(P)) ** 2)
    )
    Ctray = (193.04 + 22.72 * d + 60.38 * d ** 2) * FBM * fq * Nact
    Cc = CVe + Ctray
    return Cc


def Exchanger_function(A, P, FM):
    K = np.array([4.3247, -0.3030, 0.163])
    log_CP = K[0] + K[1] * np.log10(A) + K[2] * (np.log10(A)) ** 2

    CP = 10 ** log_CP

    C = np.array([0.03881, -0.11272, 0.08183])

    log_FP = C[0] + C[1] * np.log10(P) + C[2] * (np.log10(P)) ** 2
    FP = 10 ** log_FP
    B1 = 1.63
    B2 = 1.66
    F_BM = B1 + B2 * FM * FP
    cbm = CP * F_BM
    return cbm


def Compressor_Cost(P, W, FM):
    # Centrifugal compressor
    K = np.array([2.2897, 1.3604, -0.1027])
    log_CP = K[0] + K[1] * np.log10(W) + K[2] * (np.log10(W)) ** 2
    CP = 10 ** log_CP

    C = np.array([0, 0, 0])

    log_FP = C[0] + C[1] * np.log10(P) + C[2] * (np.log10(P)) ** 2
    FP = 10 ** log_FP
    B1 = 1.63
    B2 = 1.66
    F_BM = B1 + B2 * FM * FP
    cbm = CP * F_BM
    return cbm


def Expander_cost(W, P, FM):
    # Axial gas turbines
    K = np.array([2.7051, 1.4398, -0.1776])
    log_CP = K[0] + K[1] * np.log10(W) + K[2] * (np.log10(W)) ** 2
    CP = 10 ** log_CP

    C = np.array([0, 0, 0])

    log_FP = C[0] + C[1] * np.log10(P) + C[2] * (np.log10(P)) ** 2
    FP = 10 ** log_FP
    B1 = 1.63
    B2 = 1.66
    F_BM = B1 + B2 * FM * FP
    cbm = CP * F_BM
    return cbm


def Towers_cost(V, P, FM):
    K = np.array([3.4974, 0.4485, 0.1074])
    log_CP = K[0] + K[1] * np.log10(V) + K[2] * (np.log10(V)) ** 2
    CP = 10 ** log_CP

    C = np.array([0, 0, 0])

    log_FP = C[0] + C[1] * np.log10(P) + C[2] * (np.log10(P)) ** 2
    FP = 10 ** log_FP
    B1 = 1.63
    B2 = 1.66
    F_BM = B1 + B2 * FM * FP
    cbm = CP * F_BM
    return cbm


def Fp_tray(nT):
    if nT < 20:
        Fp = 10 ** (0.471 + 0.08516 * np.log10(nT) - 0.3473 * np.log(nT) ** 2)
    else:
        Fp = 1
    return Fp


def Tray_CP(text, area):
    K_dic = {
        "Sieve": np.array([2.9949, 0.4465, 0.3961]),
        "Valve": np.array([3.3322, 0.4838, 0.3434]),
        "Demister": np.array([3.4974, 0.4838, 0.3434]),
    }
    K = K_dic[text]
    log_CP = K[0] + K[1] * np.log10(area) + K[2] * (np.log10(area)) ** 2
    CP = 10 ** log_CP
    return CP


def Tray_cost(Area, nT, FM, type_tray):
    CP = Tray_CP(type_tray, Area)
    FP = Fp_tray(nT)
    B1 = 1.63
    B2 = 1.66
    F_BM = B1 + B2 * FM * FP
    cbm = CP * F_BM
    cbm = cbm * nT  # All the trays
    return cbm


def MSHE_COST(Volume, Pressure):
    # Pressure must be at bar
    # Volume must be at m**3
    if Pressure < 25 and 0 < Pressure:
        FP = 1
    if Pressure < 40 and 25 < Pressure:
        FP = 1.1
    if Pressure < 60 and 40 < Pressure:
        FP = 1.15
    if Pressure < 80 and 60 < Pressure:
        FP = 1.25
    if Pressure >= 80:
        FP = 1.5

    if Volume < 0.1:
        Cost = FP * 24965 * Volume ** (-0.872)
    if Volume < 1 and 0.1 < Volume:
        Cost = FP * 45082 * Volume ** (-0.645)
    if Volume >= 1:
        Cost = FP * 45598 * Volume ** (-0.535)

    return Cost


