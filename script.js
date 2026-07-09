const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;

const startBtn = document.querySelector("#startBtn");
const voiceLabel = document.querySelector("#voiceLabel");
const transcriptText = document.querySelector("#transcriptText");
const liveText = document.querySelector("#liveText");

let recognition;
let wantsListening = false;
let isListening = false;
let finalPhrases = [];

const englishFixes = [
  ["თაიფ შით", "type shit"],
  ["თაიფშით", "type shit"],
  ["ტაიფ შით", "type shit"],
  ["ტაიფშით", "type shit"],
  ["თაიფ shit", "type shit"],
  ["ტაიფ shit", "type shit"],
  ["ო მაი გად", "oh my god"],
  ["ომაიგად", "oh my god"],
  ["ბრო", "bro"],
  ["ბროუ", "bro"],
  ["ფაქ", "fuck"],
  ["შით", "shit"],
  ["დემნ", "damn"],
  ["დამნ", "damn"],
  ["ლოლ", "lol"],
  ["ოქეი", "okay"],
  ["ოკეი", "okay"],
  ["ფაინ", "fine"],
  ["ნაის", "nice"],
  ["ქულ", "cool"],
  ["ქრინჯ", "cringe"],
  ["ქრინჯი", "cringe"],
  ["სლეი", "slay"],
  ["სლეიი", "slay"],
  ["გეიმ", "game"],
  ["გეიმი", "game"],
  ["გეიმინგ", "gaming"],
  ["ჩეთი", "chat"],
  ["ჩატ", "chat"],
  ["ლაიქ", "like"],
  ["ლაიქი", "like"],
  ["დისლაიქ", "dislike"],
  ["საბსქრაიბ", "subscribe"],
  ["ფოლოუ", "follow"],
  ["ინტერნეტ", "internet"],
  ["ვაიფაი", "wifi"],
  ["აიფონ", "iPhone"],
  ["აიფონი", "iPhone"],
  ["იუთუბ", "YouTube"],
  ["იუთუბი", "YouTube"],
  ["ტიკტოკ", "TikTok"],
  ["ინსტაგრამ", "Instagram"],
  ["ინსტაგრამი", "Instagram"],
];

function cleanText(text) {
  return fixEnglishTerms(text.replace(/\s+/g, " ").trim());
}

function fixEnglishTerms(text) {
  let fixedText = text;

  for (const [georgian, english] of englishFixes) {
    fixedText = fixedText.replaceAll(georgian, english);
  }

  return fixedText;
}

function renderTranscript() {
  const text = finalPhrases.slice(-1).join(" ").trim();
  transcriptText.textContent = text || "აქ დაიწერება შენი ნალაპარაკები";
}

function setListeningState(listening) {
  isListening = listening;
  startBtn.classList.toggle("listening", listening);
  startBtn.setAttribute("aria-label", listening ? "მოსმენის შეჩერება" : "მოსმენის დაწყება");
  voiceLabel.textContent = listening ? "გისმენ..." : "დააჭირე და ილაპარაკე";
}

function startListening() {
  if (!recognition) {
    return;
  }

  wantsListening = true;

  try {
    recognition.start();
  } catch (error) {
    setListeningState(false);
  }
}

function stopListening() {
  wantsListening = false;

  if (recognition && isListening) {
    recognition.stop();
    return;
  }

  setListeningState(false);
}

if (!SpeechRecognition) {
  startBtn.disabled = true;
  voiceLabel.textContent = "გახსენი Chrome-ში ან Edge-ში";
} else {
  recognition = new SpeechRecognition();
  recognition.lang = "ka-GE";
  recognition.continuous = true;
  recognition.interimResults = true;
  recognition.maxAlternatives = 1;

  recognition.addEventListener("start", () => {
    setListeningState(true);
  });

  recognition.addEventListener("end", () => {
    setListeningState(false);

    if (wantsListening) {
      setTimeout(() => {
        try {
          recognition.start();
        } catch (error) {
          wantsListening = false;
        }
      }, 350);
    }
  });

  recognition.addEventListener("error", (event) => {
    if (event.error === "not-allowed" || event.error === "audio-capture") {
      wantsListening = false;
      setListeningState(false);
      voiceLabel.textContent = "მიკროფონი ჩართე";
    }
  });

  recognition.addEventListener("result", (event) => {
    let currentText = "";

    for (let i = event.resultIndex; i < event.results.length; i += 1) {
      const phrase = cleanText(event.results[i][0].transcript);

      if (!phrase) {
        continue;
      }

      if (event.results[i].isFinal) {
        finalPhrases.push(phrase);
        finalPhrases = finalPhrases.slice(-8);
      } else {
        currentText = `${currentText} ${phrase}`.trim();
      }
    }

    renderTranscript();
    liveText.textContent = currentText;
  });
}

startBtn.addEventListener("click", () => {
  if (wantsListening) {
    stopListening();
    return;
  }

  startListening();
});
