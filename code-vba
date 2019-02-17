Public Class Form1

    Private Sub btnRecalc_Click(sender As Object, e As EventArgs) Handles btnRecalc.Click
        Try
            Static decX3, decY3 As Decimal

            'get coordinates from user input
            decX3 = CDec(txtX3.Text)
            decY3 = CDec(txtY3.Text)
            lblX1.Text = "0,00"
            lblY1.Text = "0,00"
            lblX2.Text = "0,00"
            lblY2.Text = decY3.ToString("n2")
            lblX3.Text = decX3.ToString("n2")
            lblY3.Text = decY3.ToString("n2")
            lblX4.Text = decX3.ToString("n2")
            lblY4.Text = "0,00"
        Catch
            MessageBox.Show("Please enter valid coordinates and try again.", "Error")
        End Try

    End Sub

    Private Sub btnRotate_Click(sender As Object, e As EventArgs) Handles btnRotate.Click
        Try
            Dim decAngle, decAngleRadian, decX1, decY1, decX2, decY2, decX3, decY3, decX4, decY4, decCenterX, decCenterY,
                decRotateX1, decRotateY1, decRotateX2, decRotateY2, decRotateX3, decRotateY3, decRotateX4, decRotateY4 As Decimal

            'store coordinates as decimals
            decX1 = CDec(lblX1.Text)
            decY1 = CDec(lblY1.Text)
            decX2 = CDec(lblX2.Text)
            decY2 = CDec(lblY2.Text)
            decX3 = CDec(lblX3.Text)
            decY3 = CDec(lblY3.Text)
            decX4 = CDec(lblX4.Text)
            decY4 = CDec(lblY4.Text)

            'get angle from user input and convert it to radians
            decAngle = CDec(txtAngle.Text)
            decAngleRadian = CDec(decAngle * Math.PI / 180)

            'find centerpoint for rotation
            decCenterX = (decX1 + decX2 + decX3 + decX4) / 4
            decCenterY = (decY1 + decY2 + decY3 + decY4) / 4

            'perform rotation of rectangle
            decRotateX1 = CDec(((decX1 - decCenterX) * Math.Cos(decAngleRadian) + (decY1 - decCenterY) * Math.Sin(decAngleRadian)) + decCenterX)
            decRotateY1 = CDec(((decX1 - decCenterX) * -Math.Sin(decAngleRadian) + (decY1 - decCenterY) * Math.Cos(decAngleRadian)) + decCenterY)
            decRotateX2 = CDec(((decX2 - decCenterX) * Math.Cos(decAngleRadian) + (decY2 - decCenterY) * Math.Sin(decAngleRadian)) + decCenterX)
            decRotateY2 = CDec(((decX2 - decCenterX) * -Math.Sin(decAngleRadian) + (decY2 - decCenterY) * Math.Cos(decAngleRadian)) + decCenterY)
            decRotateX3 = CDec(((decX3 - decCenterX) * Math.Cos(decAngleRadian) + (decY3 - decCenterY) * Math.Sin(decAngleRadian)) + decCenterX)
            decRotateY3 = CDec(((decX3 - decCenterX) * -Math.Sin(decAngleRadian) + (decY3 - decCenterY) * Math.Cos(decAngleRadian)) + decCenterY)
            decRotateX4 = CDec(((decX4 - decCenterX) * Math.Cos(decAngleRadian) + (decY4 - decCenterY) * Math.Sin(decAngleRadian)) + decCenterX)
            decRotateY4 = CDec(((decX4 - decCenterX) * -Math.Sin(decAngleRadian) + (decY4 - decCenterY) * Math.Cos(decAngleRadian)) + decCenterY)

            'display new coordinates
            lblX1.Text = decRotateX1.ToString("n2")
            lblY1.Text = decRotateY1.ToString("n2")
            lblX2.Text = decRotateX2.ToString("n2")
            lblY2.Text = decRotateY2.ToString("n2")
            lblX3.Text = decRotateX3.ToString("n2")
            lblY3.Text = decRotateY3.ToString("n2")
            lblX4.Text = decRotateX4.ToString("n2")
            lblY4.Text = decRotateY4.ToString("n2")
        Catch
            MessageBox.Show("Please enter a valid rotation angle and try again.", "Error")
        End Try

    End Sub


    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Close()
    End Sub


End Class
