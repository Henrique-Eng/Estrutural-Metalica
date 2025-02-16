Public Class Class1
    Function Nc_Rd_PerfilU_kN(A_cm2, bw_mm, bf_mm, D_mm, t_mm, ri_mm, Ix_cm4, Wx_cm3, rx_cm, xg_cm, x0_cm, Iy_cm4, Wy_cm3, ry_cm, yg_cm, y0_cm, Ixy_cm4, I1_cm4, I2_cm4, alfap_graus, It_cm4, Cw_cm6, r0_cm, Kx, Ky, Kz, Lx_cm, Ly_cm, Lz_cm, E_GPa, fy_MPa)

        Dim pi As Double
        pi = 3.14159265358979

        Dim u As Double
        u = 0.3

        G_Gpa = E_GPa / (2 * (1 + u))


        ''1 - Flambagem global elástica por flexão e torção''
        'Nex(kN)'
        Dim Nex As Double
        Nex = (pi ^ 2 * E_GPa * Ix_cm4) / ((Kx * Lx_cm) ^ 2) * 100
        'Ney(kN)'
        Dim Ney As Double
        Ney = (pi ^ 2 * E_GPa * Iy_cm4) / ((Ky * Ly_cm) ^ 2) * 100
        'Nez(kN)'
        Dim Nez As Double
        Nez = (1 / (r0_cm ^ 2) * ((pi ^ 2 * E_GPa * Cw_cm6) / ((Kz * Lz_cm) ^ 2) + G_Gpa * It_cm4)) * 100
        'Nexz(kN)'
        Dim Nexz As Double
        Nexz = ((Nex + Nez) / (2 * (1 - (x0_cm / r0_cm) ^ 2)) * (1 - (1 - (4 * Nex * Nez * (1 - (x0_cm / r0_cm) ^ 2)) / ((Nex + Nez) ^ 2)) ^ 0.5))

        'Mínimo entre Nex, Ney e Nexz'
        Dim Ne As Double
        Ne = WorksheetFunction.Min(Nex, Ney, Nexz)

        ''2 -  Propriedades Geometríca de flambage''
        'lambda 0'
        Dim lambda0 As Double
        lambda0 = (((A_cm2 / 100 ^ 2) * fy_MPa * 1000) / Ne) ^ 0.5
        Dim X As Double
        'X condicional'
        If lambda0 <= 1.5 Then
            X = 0.658 ^ (lambda0 ^ 2)
        ElseIf lambda0 > 1.5 Then
            X = 0.877 / (lambda0 ^ 2)
        End If
        'Sigma'
        Dim sigma_MPa As Double
        sigma_MPa = X * fy_MPa

        ''3 - Método das larguras efetivas''
        '3.1 Mesa'
        'bmesa(mm)'
        Dim bmesa_mm As Double
        bmesa_mm = bf_mm - 1 * ri_mm
        'k mesa'
        Dim k_mesa As Double
        k_mesa = 0.43
        'lambda mesa'
        Dim lambda_mesa As Double
        lambda_mesa = (bmesa_mm / t_mm) / (0.95 * (k_mesa * (E_GPa * 1000) / sigma_MPa) ^ 0.5)
        'Largura efetiva da mesa'
        Dim b_mesa_ef As Double
        If lambda_mesa < 0.673 Then
            b_mesa_ef = bmesa_mm
        ElseIf lambda_mesa >= 0.673 Then
            b_mesa_ef = (bmesa_mm * (1 - 0.22 / lambda_mesa)) / lambda_mesa
        End If
        '3.2 Alma
        'b_alma(mm)'
        Dim b_alma_mm As Double
        b_alma_mm = bw_mm - 2 * ri_mm
        'k alma'
        Dim k_alma As Double
        k_alma = 4
        'lambda alma'
        Dim lambda_alma As Double
        lambda_alma = (b_alma_mm / t_mm) / (0.95 * (k_alma * (E_GPa * 1000) / sigma_MPa) ^ 0.5)
        'Largura efetiva da alma'
        Dim b_alma_ef As Double
        If lambda_alma < 0.673 Then
            b_alma_ef = b_alma_mm
        ElseIf lambda_alma >= 0.673 Then
            b_alma_ef = (b_alma_mm * (1 - 0.22 / lambda_alma)) / lambda_alma
        End If

        '4 - Verificação da resistência à compressão'
        'Área efetiva(mm2)'
        Dim Aef_cm2 As Double
        Aef_cm2 = A_cm2 - 2 * (bmesa_mm - b_mesa_ef) * t_mm - (b_alma_mm - b_alma_ef) * t_mm
        'Coeficiente de segurança'
        Dim gama As Double
        gama = 1.2
        'Resistência à compressão'
        Nc_Rd_PerfilU_kN = (X * Aef_cm2 * fy_MPa) / gama / 10

    End Function

End Class