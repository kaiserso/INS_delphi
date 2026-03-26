#!/usr/bin/env python3
"""
generate_guia_tb.py — Generate TB Delphi W1 Guia Rápido HTML with embedded QR code.

Usage:
    python3 generate_guia_tb.py <qr_image_path> [output_path]

Examples:
    python3 generate_guia_tb.py tb_qr.png
    python3 generate_guia_tb.py tb_qr.png output/guia_tb.html
"""

import sys
import base64
import mimetypes
from pathlib import Path


def embed_image(image_path: str) -> str:
    """Read an image file and return a base64 data URI."""
    path = Path(image_path)
    if not path.exists():
        raise FileNotFoundError(f"QR image not found: {image_path}")
    mime, _ = mimetypes.guess_type(str(path))
    if mime is None:
        mime = "image/png"
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    return f"data:{mime};base64,{b64}"


def build_html(qr_data_uri: str) -> str:
    return f"""<!DOCTYPE html>
<html lang="pt">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Guia Rápido — Oficina Delphi W1 · TB/Tuberculose</title>
<style>
  /* ── Reset & base ── */
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

  :root {{
    --blue-dark:    #1a2b3c;
    --blue-sec:     #2564a8;
    --blue-link:    #2564a8;
    --blue-italic:  #000758;
    --green-sec:    #1e7e34;
    --red-sec:      #c03929;
    --orange-sec:   #e67e22;
    --navy-text:    #1a1a1a;
    --grey-alt:     #f5f5f0;
    --cream:        #fef9e7;
    --white:        #ffffff;
    --footer-blue:  #00066f;
    --divider:      #000999;
    font-size: 8.5pt;
    font-family: Helvetica, Arial, sans-serif;
    color: var(--navy-text);
  }}

  body {{
    background: #e8e8e8;
    display: flex;
    justify-content: center;
    padding: 20px;
  }}

  .page {{
    width: 210mm;
    min-height: 297mm;
    background: #fff;
    padding: 10mm 10mm 8mm 10mm;
    display: flex;
    flex-direction: column;
    box-shadow: 0 2px 12px rgba(0,0,0,.25);
  }}

  .title-bar {{
    background: var(--blue-dark);
    color: var(--white);
    padding: 5px 6px 4px;
    margin-bottom: 6px;
    border-radius: 2px;
  }}
  .title-bar h1 {{ font-size: 11pt; font-weight: bold; line-height: 1.2; }}
  .title-bar .subtitle {{ font-size: 6pt; opacity: .85; margin-top: 2px; }}

  .columns {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6px;
    flex: 1;
  }}

  .sec-bar {{
    border-radius: 3px;
    color: var(--white);
    font-size: 7.5pt;
    font-weight: bold;
    padding: 3px 7px;
    margin-bottom: 5px;
  }}
  .sec-bar.blue   {{ background: var(--blue-sec); }}
  .sec-bar.green  {{ background: var(--green-sec); }}
  .sec-bar.red    {{ background: var(--red-sec); }}
  .sec-bar.orange {{ background: var(--orange-sec); }}

  .qr-block {{
    display: flex;
    flex-direction: column;
    gap: 6px;
    margin-bottom: 8px;
    align-items: flex-start;
  }}
  .qr-block img {{
    width: 232px;
    height: 232px;
    flex-shrink: 0;
    image-rendering: pixelated;
  }}
  .qr-text {{ font-size: 7pt; line-height: 1.4; }}
  .qr-text .link {{ color: var(--blue-link); font-weight: bold; font-size: 7pt; }}
  .qr-text .hint {{ color: var(--blue-italic); font-style: italic; font-size: 6.5pt; }}

  .step {{ display: flex; gap: 7px; margin-bottom: 7px; align-items: flex-start; }}
  .step-num {{
    flex-shrink: 0;
    width: 15px; height: 15px;
    border-radius: 50%;
    background: var(--blue-sec);
    color: var(--white);
    font-size: 7pt; font-weight: bold;
    display: flex; align-items: center; justify-content: center;
    margin-top: 1px;
  }}
  .step-body {{ flex: 1; font-size: 7pt; line-height: 1.4; }}
  .step-body .step-title {{
    color: var(--blue-sec); font-weight: bold; font-size: 7pt;
    display: block; margin-bottom: 2px;
  }}
  .step-body p {{ color: var(--navy-text); }}

  .q-table {{ width: 100%; border-collapse: collapse; font-size: 7pt; margin-bottom: 4px; }}
  .q-table tr:nth-child(odd)  td {{ background: var(--grey-alt); }}
  .q-table tr:nth-child(even) td {{ background: var(--white); }}
  .q-table td {{ padding: 2.5px 3px; vertical-align: top; line-height: 1.35; }}
  .q-table .code  {{ font-weight: bold; width: 28px; white-space: nowrap; }}
  .q-table .label {{ font-weight: bold; width: 110px; }}
  .q-table .detail{{ color: var(--blue-italic); font-style: italic; }}

  .salto {{
    background: var(--grey-alt);
    padding: 4px 5px; font-size: 7pt;
    border-radius: 2px; margin-top: 2px;
  }}
  .salto strong {{ display: block; margin-bottom: 2px; }}

  .bullet-list {{ list-style: none; font-size: 7pt; line-height: 1.4; }}
  .bullet-list li {{ display: flex; gap: 5px; margin-bottom: 5px; align-items: flex-start; }}
  .bullet-list li .bul {{ flex-shrink: 0; margin-top: 1px; }}
  .bullet-list.dark .bul {{ color: var(--navy-text); font-size: 8pt; }}

  .prob-hdr {{
    font-weight: bold; color: var(--red-sec); font-size: 7pt;
    display: flex; gap: 5px; align-items: flex-start; margin-bottom: 2px;
  }}
  .prob-hdr .bul {{ color: var(--red-sec); font-size: 8pt; flex-shrink: 0; }}
  .prob-body {{
    font-size: 7pt; color: var(--navy-text);
    margin-bottom: 7px; padding-left: 12px; line-height: 1.35;
  }}

  .plano-b {{
    background: var(--cream);
    padding: 5px 6px; font-size: 7pt; border-radius: 2px;
  }}
  .plano-b ol {{ list-style: none; }}
  .plano-b ol li {{
    display: flex; gap: 7px;
    padding: 3px 0; line-height: 1.4; align-items: flex-start;
  }}
  .plano-b ol li .n {{ font-weight: bold; flex-shrink: 0; width: 10px; }}

  .footer {{
    border-top: 0.5px solid var(--divider);
    margin-top: 6px; padding-top: 3px;
    text-align: center; font-size: 6pt; color: var(--footer-blue);
  }}

  @media print {{
    body {{ background: none; padding: 0; }}
    .page {{ box-shadow: none; }}
  }}
</style>
</head>
<body>
<div class="page">

  <div class="title-bar">
    <h1>Guia Rápido — Oficina Delphi W1 · TB/Tuberculose</h1>
    <div class="subtitle">Instruções de acesso e utilização do KoboToolbox | Confidencial — uso exclusivo dos participantes</div>
  </div>

  <div class="columns">

    <!-- LEFT COLUMN -->
    <div class="left-col">

      <div class="sec-bar blue">① ACESSO — PRIMEIRO PASSO</div>

      <div class="qr-block">
        <img src="{qr_data_uri}" alt="QR Code">
        <div class="qr-text">
          <strong>Ligue o seu dispositivo ao WiFi<br>
          e abra o link abaixo ou<br>
          aponte a câmara para o QR:</strong><br>
          <span class="link">delphi-w1.pages.dev/tb/gateway.html</span><br>
          <span class="hint">(ou escreva directamente no browser)</span>
        </div>
      </div>

      <div class="step">
        <div class="step-num">1</div>
        <div class="step-body">
          <span class="step-title">Introduza o seu email e receba o código</span>
          <p>Na página de entrada, escreva o email com que se registou e clique em 'Enviar código'. Receberá um código de 6 dígitos por email. Verifique também a pasta de spam/lixo.</p>
        </div>
      </div>

      <div class="step">
        <div class="step-num">2</div>
        <div class="step-body">
          <span class="step-title">Introduza o código de 6 dígitos</span>
          <p>Copie ou escreva o código recebido no campo indicado e clique em 'Verificar'. O acesso é válido durante toda a sessão — não precisa de repetir se reabrir o link.</p>
        </div>
      </div>

      <div class="step">
        <div class="step-num">3</div>
        <div class="step-body">
          <span class="step-title">Seleccione o seu código de especialista</span>
          <p>Após autenticação verá o painel principal. Seleccione o seu Código de Especialista (ex: 001-TB) e o sub-formulário correcto. Use o MESMO código em todas as oficinas.</p>
        </div>
      </div>

      <div class="step">
        <div class="step-num">4</div>
        <div class="step-body">
          <span class="step-title">Leia cada intervenção antes de responder</span>
          <p>Cada pergunta inclui um link '■ Ver ficha'. Clique para abrir os detalhes completos da intervenção numa nova janela antes de avaliar. Calcule aprox. 3–5 min por intervenção.</p>
        </div>
      </div>

      <div class="sec-bar blue" style="margin-top:4px;">② PERGUNTAS — REFERÊNCIA RÁPIDA</div>

      <table class="q-table">
        <tr><td class="code">M1</td><td class="label">Nível de expertise</td><td class="detail">1 = Geral · 2 = Intermédio · 3 = Alto</td></tr>
        <tr><td class="code">M2</td><td class="label">Necessita optimização?</td><td class="detail">Sim definitivamente / Possivelmente / Não</td></tr>
        <tr><td class="code">M3</td><td class="label">Existe duplicação?</td><td class="detail">Visível só se M2 ≠ Não</td></tr>
        <tr><td class="code">M4</td><td class="label">Quais programas duplicados?</td><td class="detail">Lista — pode seleccionar múltiplos</td></tr>
        <tr><td class="code">M5</td><td class="label">Pode ser integrada?</td><td class="detail">Visível só se M2 ≠ Não</td></tr>
        <tr><td class="code">M6</td><td class="label">Com que intervenção(ões)?</td><td class="detail">Lista — pode seleccionar múltiplos</td></tr>
        <tr><td class="code">M7</td><td class="label">Pode reduzir recursos?</td><td class="detail">Visível só se M2 ≠ Não</td></tr>
        <tr><td class="code">M8</td><td class="label">Outro motivo de optimização?</td><td class="detail">Visível só se M2 ≠ Não</td></tr>
        <tr><td class="code">M9</td><td class="label">Descreva outro motivo</td><td class="detail">Campo de texto — só se M8 = Sim</td></tr>
        <tr><td class="code">M10</td><td class="label">Impacto da optimização</td><td class="detail">1 = Baixo · 2 = Médio · 3 = Alto</td></tr>
        <tr><td class="code">M11</td><td class="label">Como optimizar?</td><td class="detail">Comentário — obrigatório se M2 ≠ Não</td></tr>
      </table>

      <div class="salto">
        <strong>SALTO AUTOMÁTICO</strong>
        Se M2 = 'Não' → o formulário oculta M3–M11 automaticamente.
      </div>

    </div><!-- /left-col -->

    <!-- RIGHT COLUMN -->
    <div class="right-col">

      <div class="sec-bar blue">③ SUBMISSÃO E CONFIRMAÇÃO</div>

      <div class="step">
        <div class="step-num">5</div>
        <div class="step-body">
          <span class="step-title">Não mude de dispositivo a meio do formulário</span>
          <p>O progresso é guardado localmente no browser. Se fechar por acidente, reabra o mesmo link no mesmo browser — pode recuperar o rascunho. Mudar de dispositivo ou browser apaga o progresso.</p>
        </div>
      </div>

      <div class="step">
        <div class="step-num">6</div>
        <div class="step-body">
          <span class="step-title">Reveja e submeta</span>
          <p>No final clique em 'Próximo' para rever. Corrija se necessário e clique em 'Enviar'. Aguarde: ✔ 'A sua resposta foi guardada'. Só aparece esta mensagem = recebido.</p>
        </div>
      </div>

      <div class="step">
        <div class="step-num">7</div>
        <div class="step-body">
          <span class="step-title">Não é possível editar após submissão</span>
          <p>Se precisar corrigir, contacte o facilitador. A equipa pode anular e reabrir para que submeta novamente.</p>
        </div>
      </div>

      <div class="sec-bar green" style="margin-top:6px;">④ BOAS PRÁTICAS</div>

      <ul class="bullet-list dark">
        <li><span class="bul">■</span><span>Não partilhe o seu código de especialista com outros participantes.</span></li>
        <li><span class="bul">■</span><span>Se fechar por acidente, reabra o mesmo link no mesmo browser — o rascunho pode ser recuperado. Não mude de dispositivo.</span></li>
        <li><span class="bul">■</span><span>Consulte a ficha da intervenção (link em cada pergunta) antes de responder.</span></li>
      </ul>

      <div class="sec-bar red" style="margin-top:6px;">⑤ PROBLEMAS E SOLUÇÕES</div>

      <div class="prob-hdr"><span class="bul">■</span><span>Link não abre / página em branco</span></div>
      <div class="prob-body">Verifique a ligação WiFi. Tente recarregar (F5). Tente noutro browser (Chrome, Firefox, Safari).</div>

      <div class="prob-hdr"><span class="bul">■</span><span>Não recebi o código de email</span></div>
      <div class="prob-body">Verifique spam/lixo. Confirme que usou o email com que se registou — um endereço diferente não funcionará. Aguarde 1–2 min ou peça ao facilitador para reenviar.</div>

      <div class="prob-hdr"><span class="bul">■</span><span>Código de email inválido ou expirou</span></div>
      <div class="prob-body">Os códigos expiram ao fim de 10 minutos. Clique em 'Reenviar' para obter um novo código.</div>

      <div class="prob-hdr"><span class="bul">■</span><span>Esqueci o meu código de especialista</span></div>
      <div class="prob-body">Consulte o cartão de identificação distribuído. Formato: 3 números + -TB (ex: 001-TB).</div>

      <div class="prob-hdr"><span class="bul">■</span><span>Não consigo submeter / campos vermelhos</span></div>
      <div class="prob-body" style="margin-bottom:4px;">M1 e M2 são sempre obrigatórios. Verifique todos os campos assinalados a vermelho.</div>

      <div class="sec-bar orange" style="margin-top:6px;">⑥ PLANO B — FICHEIRO EXCEL (SE KOBO FALHAR)</div>

      <div class="plano-b">
        <ol>
          <li><span class="n">1</span><span>Receba o ficheiro do facilitador: delphi_w1_tb_[CODIGO].xlsx</span></li>
          <li><span class="n">2</span><span>Folha 'Respostas': seleccione o seu código na célula amarela no topo</span></li>
          <li><span class="n">3</span><span>Preencha uma linha por intervenção (M1 e M2 sempre obrigatórios)</span></li>
          <li><span class="n">4</span><span>Se M2 ≠ Não: preencha M3–M11. Em M4/M6 escreva os códigos separados por vírgula (ex: tb_03, tb_07)</span></li>
          <li><span class="n">5</span><span>Guarde e entregue o ficheiro ao facilitador no final da sessão</span></li>
        </ol>
      </div>

    </div><!-- /right-col -->
  </div><!-- /columns -->

  <div class="footer">
    Delphi W1 · TB/Tuberculose · 2026 | Confidencial — uso exclusivo dos participantes
  </div>

</div><!-- /page -->
</body>
</html>"""


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    qr_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "delphi_w1_tb_guia_rapido.html"

    try:
        qr_data_uri = embed_image(qr_path)
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    html = build_html(qr_data_uri)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Generated: {output_path}")


if __name__ == "__main__":
    main()
