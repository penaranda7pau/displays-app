"""Genera la miniatura (thumb_b64) de los reportes viejos que solo tienen foto completa.

Uso (DATABASE_URL debe apuntar a la base que corresponda; NO correr sobre producción sin aprobación):
    python generar_miniaturas.py            # procesa todo, en lotes
    python generar_miniaturas.py --limite 50  # solo los primeros 50 (prueba)
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import app as A  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limite", type=int, default=0)
    ap.add_argument("--lote", type=int, default=50)
    args = ap.parse_args()

    total = sin_miniatura = 0
    with A.app.app_context():
        # 1) fotos completas de reportes; 2) fotos de góndola (FotoGondola)
        for modelo, campo_foto in ((A.Reporte, "foto_b64"), (A.FotoGondola, "foto_b64")):
            procesados = 0
            while True:
                ids = [i for (i,) in (A.db.session.query(modelo.id)
                                      .filter(modelo.thumb_b64.is_(None), getattr(modelo, campo_foto).isnot(None),
                                              getattr(modelo, campo_foto) != "")
                                      .order_by(modelo.id).limit(args.lote).all())]
                if not ids:
                    break
                for rid in ids:
                    obj = A.db.session.get(modelo, rid)
                    miniatura = A._miniatura_b64(getattr(obj, campo_foto))
                    if not miniatura:
                        sin_miniatura += 1
                    obj.thumb_b64 = miniatura or ""   # "" = no se pudo; no se reintenta (la lista muestra 📷)
                    procesados += 1
                    if args.limite and procesados >= args.limite:
                        break
                A.db.session.commit()
                A.db.session.expunge_all()
                print(f"{modelo.__name__}: {procesados}", flush=True)
                if args.limite and procesados >= args.limite:
                    break
            total += procesados
    print(f"listo. procesados={total} sin_miniatura={sin_miniatura}")


if __name__ == "__main__":
    main()
