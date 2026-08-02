"""
Moteur de génération du fichier "Performance Drive" à partir du fichier
Preparation hebdomadaire et (optionnellement) d'un planning PDF. Expose les
endpoints HTTP que le workflow n8n appelle.

Contrat de l'API :

POST /generate?taux_actuelle=0.0245&taux_precedente=0.0219&productivite_s1={"MONTESCOT1":12.5}
Corps de la requête : multipart/form-data avec deux champs fichier :
  - preparation : le fichier Preparation .xlsx (obligatoire)
  - planning    : le planning PDF hebdomadaire (optionnel — si absent, la
                  colonne "Heure travaillée" reste manuelle comme avant)
Réponse : le fichier "Performance Drive S{semaine}.xlsx" généré, en pièce
jointe (Content-Disposition).

POST /productivity : mêmes paramètres que /generate, mais renvoie un JSON
simple (pas de fichier) avec la semaine détectée et la productivité calculée
par employé cette semaine : {"detected_week": 31, "employee_productivity":
{"MONTESCOT1": {"nom": "...", "productivite_h": 54.1}, ...}}. Ce second appel
existe pour que n8n puisse persister l'historique par employé sans avoir à
lire des en-têtes HTTP personnalisés sur une réponse fichier.

Lancement local :
    pip install -r requirements.txt
    uvicorn main:app --reload

Déploiement (Render/Railway/Fly.io) :
    Commande de démarrage : uvicorn main:app --host 0.0.0.0 --port $PORT
"""
import os
import json
import tempfile

from fastapi import FastAPI, File, UploadFile, Query, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from typing import Optional

import gen_lib

app = FastAPI(title="Moteur Performance Drive")


@app.get("/health")
def health():
    return {"status": "ok"}


def _parse_productivite_s1(productivite_s1: Optional[str]) -> dict:
    productivite_s1_map = {}
    if productivite_s1:
        try:
            parsed = json.loads(productivite_s1)
            # accepte soit {matricule: valeur}, soit {matricule: {"productivite_h": valeur}}
            for k, v in parsed.items():
                if isinstance(v, dict):
                    productivite_s1_map[k] = v.get('productivite_h')
                else:
                    productivite_s1_map[k] = v
        except Exception as e:
            print(f"Avertissement: productivite_s1 illisible ({e}), ignore.")
    return productivite_s1_map


async def _run_build(preparation: UploadFile, planning: Optional[UploadFile],
                      taux_actuelle, taux_precedente, semaine, productivite_s1):
    prep_bytes = await preparation.read()
    if not prep_bytes:
        raise HTTPException(status_code=400, detail="Fichier Preparation vide ou manquant.")

    productivite_s1_map = _parse_productivite_s1(productivite_s1)

    work_dir = tempfile.mkdtemp(prefix="perfdrive_")
    src_path = os.path.join(work_dir, "source.xlsx")
    with open(src_path, "wb") as f:
        f.write(prep_bytes)

    planning_path = None
    if planning is not None:
        planning_bytes = await planning.read()
        if planning_bytes:
            planning_path = os.path.join(work_dir, "planning.pdf")
            with open(planning_path, "wb") as f:
                f.write(planning_bytes)

    try:
        detected_week, employees = gen_lib.get_week_and_employees(src_path)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Impossible de lire le fichier Preparation envoyé : {e}")

    if not employees:
        raise HTTPException(status_code=400, detail="Aucun employé détecté dans l'onglet Preparation — vérifie le fichier source.")

    if semaine is not None and int(semaine) != detected_week:
        print(f"Avertissement: semaine indiquée dans l'email ({semaine}) != semaine déduite du fichier ({detected_week}). Semaine déduite utilisée.")

    out_name = f"Performance Drive S{detected_week}.xlsx"
    out_path = os.path.join(work_dir, out_name)

    try:
        week, n, employee_productivity = gen_lib.build(
            src_path,
            out_path,
            taux_actuelle=taux_actuelle,
            taux_precedente=taux_precedente,
            planning_path=planning_path,
            productivite_s1=productivite_s1_map,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur lors de la lecture du planning ou de la génération : {e}")

    return work_dir, out_path, out_name, week, employee_productivity


@app.post("/generate")
async def generate(
    preparation: UploadFile = File(..., description="Fichier Preparation .xlsx"),
    planning: Optional[UploadFile] = File(default=None, description="Planning PDF hebdomadaire (optionnel)"),
    taux_actuelle: Optional[float] = Query(default=None, description="Taux de rupture de la semaine courante (fraction, ex: 0.0245 pour 2.45%)"),
    taux_precedente: Optional[float] = Query(default=None, description="Taux de rupture de la semaine précédente (fraction)"),
    semaine: Optional[int] = Query(default=None, description="Numéro de semaine indiqué par l'expéditeur (informatif)"),
    productivite_s1: Optional[str] = Query(default=None, description="JSON {matricule: productivite_h} de la semaine précédente"),
):
    try:
        work_dir, out_path, out_name, week, employee_productivity = await _run_build(
            preparation, planning, taux_actuelle, taux_precedente, semaine, productivite_s1,
        )
        return FileResponse(
            out_path,
            filename=out_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur lors de la génération du fichier : {e}")
    # Note: work_dir is intentionally not cleaned up synchronously here since
    # FileResponse streams the file after this function returns.


@app.post("/productivity")
async def productivity(
    preparation: UploadFile = File(..., description="Fichier Preparation .xlsx"),
    planning: Optional[UploadFile] = File(default=None, description="Planning PDF hebdomadaire (optionnel)"),
    taux_actuelle: Optional[float] = Query(default=None),
    taux_precedente: Optional[float] = Query(default=None),
    semaine: Optional[int] = Query(default=None),
    productivite_s1: Optional[str] = Query(default=None),
):
    """Meme traitement que /generate mais renvoie uniquement un JSON (semaine
    detectee + productivite par employe), sans le fichier. Appele par n8n en
    parallele de /generate pour alimenter l'historique par employe."""
    try:
        _work_dir, _out_path, _out_name, week, employee_productivity = await _run_build(
            preparation, planning, taux_actuelle, taux_precedente, semaine, productivite_s1,
        )
        return JSONResponse({
            "detected_week": week,
            "employee_productivity": employee_productivity,
        })
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur lors du calcul de productivité : {e}")


@app.post("/podium")
async def podium(
    preparation: UploadFile = File(..., description="Fichier Preparation .xlsx"),
    planning: Optional[UploadFile] = File(default=None, description="Planning PDF hebdomadaire (optionnel)"),
    taux_actuelle: Optional[float] = Query(default=None, description="Taux de rupture de la semaine courante (fraction)"),
    taux_precedente: Optional[float] = Query(default=None, description="Taux de rupture de la semaine précédente (fraction)"),
    semaine: Optional[int] = Query(default=None),
    productivite_s1: Optional[str] = Query(default=None),
):
    """Meme traitement que /generate mais renvoie le PDF "podium" visuel
    (classement + rappel taux de rupture, charte Intermarché) destine a etre
    affiche en salle de pause. Appele par n8n en parallele de /generate."""
    try:
        work_dir, _out_path, _out_name, week, employee_productivity = await _run_build(
            preparation, planning, taux_actuelle, taux_precedente, semaine, productivite_s1,
        )
        pdf_name = f"Podium Performance Drive S{week}.pdf"
        pdf_path = os.path.join(work_dir, pdf_name)
        gen_lib.generate_podium_pdf(
            pdf_path, week, employee_productivity,
            taux_actuelle=taux_actuelle, taux_precedente=taux_precedente,
        )
        return FileResponse(
            pdf_path,
            filename=pdf_name,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{pdf_name}"'},
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur lors de la génération du podium : {e}")
