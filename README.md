# PA-Config
## Warning
execute this command to be able to run scripts (only enables it for the current powershell process):
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
## excel_install.ps1
Le script instale excel avec l'outil de microsoft 365 et après il change les settings pour ne jamais s'ouvrir en mode protéger, Tous les macro son aussi activer.
## power_automate_fix.ps1
Le script règle l'erreur de ScriptRunNotFound selon le fix de microsoft : https://admin.powerplatform.microsoft.com/support/knownissues/6476290 
Il créer aussi un service qui roule comme system qui s'occupe de supprimer manuellement les différent fichier qui s'accumule quand on disable le file cleanup qui cause l'erreur.