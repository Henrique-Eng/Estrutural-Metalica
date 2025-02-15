Imports Microsoft.VisualBasic
Imports Microsoft.Interop.Excel.Application


Public Class Class1
    Function Nc_Rd_PerfilU(A_cm2, bw_mm, bf_mm, D_mm, t_mm, ri_mm, Ix_cm4, Wx_cm3, rx_cm, xg_cm, x0_cm, Iy_cm4, Wy_cm3, ry_cm, yg_cm, y0_cm, Ixy_cm4, I1_cm4, I2_cm4, alfap_graus, It_cm4, Cw_cm6, r0_cm, Kx, Ky, Kz, Lx_cm, Ly_cm, Lz_cm, E_GPa, fy_MPa)

        Dim pi As Double
        pi = 3.14159265358979

        Dim u As Double
        u = 0.3

        G_Gpa = E_GPa / (2 * (1 + u))


        'Flambagem global elástica por flexão e torção'
        'Nex(kN)
        Nex = (pi ^ 2 * E_GPa * Ix_cm4) / ((Kx * Lx_cm) ^ 2)
        'Ney(kN)'
        Ney = (pi ^ 2 * E_GPa * Iy_cm4) / ((Ky * Ly_cm) ^ 2)
        'Nez(kN)'
        Nez = 1 / (r0_cm ^ 2) * ((pi ^ 2 * E_GPa * Cw_cm6) / ((Kz * Lz_cm) ^ 2) + G_Gpa * It_cm4)
        'Nexz(kN)'
        Nexz = ((Nex + Nez) / (2 * (1 - (x0_cm / r0_cm) ^ 2)) * (1 - (1 - (4 * Nex * Nez * (1 - (x0_cm / r0_cm) ^ 2)) / ((Nex + Nez) ^ 2)) ^ 0.5)) * 100

        'Mínimo entre Nex, Ney e Nexz'
        Ne = WorksheetFunction.Min(Nex, Ney, Nexz)



        Nc_Rd_PerfilU = Nex
    End Function

End Class