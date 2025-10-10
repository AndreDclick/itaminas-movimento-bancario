try {
    [Console]::Out.Flush()

    Start-Process "C:\Users\rpa.dclick\Documents\projetos\itaminas-conciliacao-fornecedores\dist\itaminas-movimentacao-bancaria.exe"
}
catch {
    Write-Output "Exce��o ao iniciar processo: $_"
    exit 2
}
finally {
    Write-Output "Processo finalizado (try/catch)."
    [Console]::Out.Flush()
}