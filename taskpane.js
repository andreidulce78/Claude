// ============================================================
// CONFIG — remplace par ta vraie clé API Anthropic
// ============================================================
const ANTHROPIC_API_KEY = "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";
// ============================================================

Office.onReady(() => {
  console.log("Claude Add-in chargé ✓");
});

const PROMPTS = {
  resume: "Résume cet e-mail en 3-5 phrases claires et concises en français. Va droit au but.",
  points: "Liste les points clés et les actions requises de cet e-mail en français. Format : bullet points avec emoji ✅ pour les actions.",
  reply: "Propose une réponse professionnelle et concise à cet e-mail en français. Adapte le ton à celui de l'email original.",
  translate: "Traduis cet e-mail en français de manière naturelle et professionnelle."
};

const TITLES = {
  resume: "Résumé",
  points: "Points clés & actions",
  reply: "Réponse suggérée",
  translate: "Traduction"
};

async function analyzeEmail(action) {
  const loading = document.getElementById("loading");
  const resultBox = document.getElementById("resultBox");
  const errorBox = document.getElementById("errorBox");
  const resultContent = document.getElementById("resultContent");
  const resultTitle = document.getElementById("resultTitle");

  // Reset UI
  resultBox.classList.remove("visible");
  errorBox.classList.remove("visible");
  loading.classList.add("visible");

  try {
    // Récupère le contenu de l'email ouvert
    const item = Office.context.mailbox.item;

    const subject = item.subject || "(sans objet)";
    const sender = item.from ? item.from.displayName + " <" + item.from.emailAddress + ">" : "Inconnu";

    // Récupère le corps de l'email
    item.body.getAsync(Office.CoercionType.Text, async (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        showError("Impossible de lire le contenu de l'e-mail.");
        return;
      }

      const body = result.value.trim().substring(0, 4000); // limite à 4000 chars

      const emailContext = `De : ${sender}\nObjet : ${subject}\n\nContenu :\n${body}`;

      try {
        const response = await fetch("https://api.anthropic.com/v1/messages", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01"
          },
          body: JSON.stringify({
            model: "claude-sonnet-4-20250514",
            max_tokens: 1000,
            messages: [{
              role: "user",
              content: `${PROMPTS[action]}\n\n---\n${emailContext}`
            }]
          })
        });

        if (!response.ok) {
          const err = await response.json();
          throw new Error(err.error?.message || "Erreur API");
        }

        const data = await response.json();
        const text = data.content[0].text;

        // Affiche le résultat
        loading.classList.remove("visible");
        resultTitle.textContent = TITLES[action];
        resultContent.textContent = text;
        resultBox.classList.add("visible");

      } catch (apiError) {
        showError("Erreur API Claude : " + apiError.message);
      }
    });

  } catch (err) {
    showError("Erreur : " + err.message);
  }
}

function showError(message) {
  const loading = document.getElementById("loading");
  const errorBox = document.getElementById("errorBox");
  loading.classList.remove("visible");
  errorBox.textContent = "⚠ " + message;
  errorBox.classList.add("visible");
}

function copyResult() {
  const text = document.getElementById("resultContent").textContent;
  navigator.clipboard.writeText(text).then(() => {
    const btn = document.querySelector(".copy-btn");
    btn.textContent = "Copié ✓";
    setTimeout(() => btn.textContent = "Copier", 2000);
  });
}
