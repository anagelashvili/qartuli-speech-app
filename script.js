const pptInput = document.querySelector("#pptInput");
const docInput = document.querySelector("#docInput");
const setupFileInput = document.querySelector("#setupFileInput");
const voiceSelect = document.querySelector("#voiceSelect");
const pptName = document.querySelector("#pptName");
const docName = document.querySelector("#docName");
const prevBtn = document.querySelector("#prevBtn");
const playBtn = document.querySelector("#playBtn");
const setupPlayBtn = document.querySelector("#setupPlayBtn");
const nextBtn = document.querySelector("#nextBtn");
const testVoiceBtn = document.querySelector("#testVoiceBtn");
const aresBtn = document.querySelector("#aresBtn");
const aresStatus = document.querySelector("#aresStatus");
const diagnosticText = document.querySelector("#diagnosticText");
const progressFill = document.querySelector("#progressFill");
const slideCounter = document.querySelector("#slideCounter");
const statusText = document.querySelector("#statusText");
const slideCanvas = document.querySelector("#slideCanvas");
const scriptText = document.querySelector("#scriptText");
const speech = window.speechSynthesis;
const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;

let slides = [];
let docSegments = [];
let docFullText = "";
let currentSlide = 0;
let isReading = false;
let advanceTimer;
let voices = [];
let readingParts = [];
let readingPartIndex = 0;
let isVisualFallback = false;
let activeAudio;
let activeTtsController;
let ignoreSpeechErrorsUntil = 0;
let lastBackendError = "";
let aresRecognition;
let aresEnabled = false;
let readingSession = 0;

function setStatus(message) {
  statusText.textContent = message;
}

function setDiagnostic(message) {
  diagnosticText.textContent = message;
}

function setAresStatus(message) {
  aresStatus.textContent = message;
}

function decodeXmlText(text) {
  const textarea = document.createElement("textarea");
  textarea.innerHTML = text;
  return textarea.value;
}

function containsGeorgian(text) {
  return /[\u10A0-\u10FF]/.test(text);
}

function transliterateGeorgian(text) {
  const map = {
    ა: "a",
    ბ: "b",
    გ: "g",
    დ: "d",
    ე: "e",
    ვ: "v",
    ზ: "z",
    თ: "t",
    ი: "i",
    კ: "k",
    ლ: "l",
    მ: "m",
    ნ: "n",
    ო: "o",
    პ: "p",
    ჟ: "zh",
    რ: "r",
    ს: "s",
    ტ: "t",
    უ: "u",
    ფ: "p",
    ქ: "k",
    ღ: "gh",
    ყ: "q",
    შ: "sh",
    ჩ: "ch",
    ც: "ts",
    ძ: "dz",
    წ: "ts",
    ჭ: "ch",
    ხ: "kh",
    ჯ: "j",
    ჰ: "h",
  };

  return [...text].map((char) => map[char] || char).join("");
}

function textFromXml(xmlText) {
  const xml = new DOMParser().parseFromString(xmlText, "application/xml");
  return [...xml.getElementsByTagName("*")]
    .filter((node) => node.localName === "t")
    .map((node) => node.textContent.trim())
    .filter(Boolean)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

function splitTextIntoSegments(text, count) {
  const paragraphs = text
    .split(/\n{2,}/)
    .map((part) => part.replace(/\s+/g, " ").trim())
    .filter(Boolean);

  if (!count || paragraphs.length === count) {
    return paragraphs;
  }

  const sentences = text
    .replace(/\s+/g, " ")
    .split(/(?<=[.!?؟։…])\s+/)
    .map((part) => part.trim())
    .filter(Boolean);

  if (sentences.length <= count) {
    return sentences;
  }

  const bucketSize = Math.ceil(sentences.length / count);
  return Array.from({ length: count }, (_, index) =>
    sentences.slice(index * bucketSize, (index + 1) * bucketSize).join(" ")
  ).filter(Boolean);
}

async function readZipFile(file) {
  if (!window.JSZip) {
    throw new Error("JSZip ვერ ჩაიტვირთა. შეამოწმე ინტერნეტი და თავიდან სცადე.");
  }

  const buffer = await file.arrayBuffer();
  return window.JSZip.loadAsync(buffer);
}

async function loadPowerPoint(file) {
  const zip = await readZipFile(file);
  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => Number(a.match(/slide(\d+)\.xml/)[1]) - Number(b.match(/slide(\d+)\.xml/)[1]));

  const parsedSlides = [];

  for (const name of slideFiles) {
    const xml = await zip.file(name).async("text");
    const text = textFromXml(xml);
    parsedSlides.push({
      title: text.split(/[.!?۔։]/)[0]?.trim() || `სლაიდი ${parsedSlides.length + 1}`,
      text: text || "ამ სლაიდზე ტექსტი ვერ მოიძებნა.",
    });
  }

  slides = parsedSlides;
  if (docFullText) {
    docSegments = splitTextIntoSegments(docFullText, slides.length);
  }
  currentSlide = 0;
  setStatus(`${slides.length} სლაიდი ჩაიტვირთა`);
  renderSlide();
}

function createSlidesFromDocSegments() {
  slides = docSegments.map((segment, index) => ({
    title: `ნაწილი ${index + 1}`,
    text: segment,
  }));
  currentSlide = 0;
}

async function loadWord(file) {
  const zip = await readZipFile(file);
  const wordXmlFiles = Object.keys(zip.files)
    .filter((name) => /^word\/.+\.xml$/.test(name) && !name.includes("_rels/"))
    .sort((a, b) => {
      if (a === "word/document.xml") return -1;
      if (b === "word/document.xml") return 1;
      return a.localeCompare(b);
    });

  if (!wordXmlFiles.length) {
    throw new Error("Word XML ფაილები ვერ მოიძებნა");
  }

  let paragraphs = [];

  for (const name of wordXmlFiles) {
    const xmlText = await zip.file(name).async("text");
    const xml = new DOMParser().parseFromString(xmlText, "application/xml");
    const xmlParagraphs = [...xml.getElementsByTagName("*")]
      .filter((node) => node.localName === "p")
      .map((paragraph) =>
        [...paragraph.getElementsByTagName("*")]
          .filter((node) => node.localName === "t")
          .map((node) => node.textContent)
          .join("")
          .replace(/\s+/g, " ")
          .trim()
      )
      .filter(Boolean);

    paragraphs.push(...xmlParagraphs);

    if (!xmlParagraphs.length) {
      const textMatches = [...xmlText.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)];
      paragraphs.push(
        ...textMatches
          .map((match) => decodeXmlText(match[1]).replace(/\s+/g, " ").trim())
          .filter(Boolean)
      );
    }
  }

  docFullText = paragraphs.join("\n\n");
  docSegments = splitTextIntoSegments(docFullText, slides.length);
  if (!slides.length) {
    createSlidesFromDocSegments();
  }
  setStatus(`Word ტექსტი ჩაიტვირთა: ${docSegments.length} ნაწილი`);
  if (!docSegments.length) {
    setStatus("Word ჩაიტვირთა, მაგრამ ტექსტი ვერ ამოვიკითხე. სცადე .docx ფორმატად ხელახლა Save As.");
  }
  renderSlide();
}

function getSlideScript(index) {
  return docSegments[index] || slides[index]?.text || "";
}

function renderSlide() {
  const slide = slides[currentSlide];
  const count = slides.length;

  slideCounter.textContent = `სლაიდი ${count ? currentSlide + 1 : 0} / ${count}`;
  progressFill.style.width = count ? `${((currentSlide + 1) / count) * 100}%` : "0%";
  prevBtn.disabled = !count || currentSlide === 0;
  nextBtn.disabled = !count || currentSlide === count - 1;
  playBtn.disabled = !count;
  setupPlayBtn.disabled = !count;

  if (!slide) {
    slideCanvas.innerHTML = `
      <p class="slideKicker">მომზადება</p>
      <h2>დაიწყე ფაილების ატვირთვით</h2>
      <p>ატვირთე .pptx პრეზენტაცია და სურვილის შემთხვევაში .docx ტექსტი. შემდეგ დააჭირე „წაკითხვა“-ს.</p>
    `;
    scriptText.textContent = "აქ გამოჩნდება ტექსტი, რომელსაც ქართული ხმა წაიკითხავს.";
    return;
  }

  const script = getSlideScript(currentSlide);
  slideCanvas.innerHTML = `
    <p class="slideKicker">ქართული ხმა</p>
    <h2>${escapeHtml(slide.title)}</h2>
    <p>${escapeHtml(slide.text)}</p>
  `;
  scriptText.textContent = script || "ამ სლაიდისთვის წასაკითხი ტექსტი არ მოიძებნა.";
}

function escapeHtml(text) {
  return text.replace(/[&<>"']/g, (char) => {
    const entities = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;",
    };
    return entities[char];
  });
}

function populateVoices() {
  if (!speech) {
    voiceSelect.innerHTML = '<option value="">ხმა მიუწვდომელია</option>';
    testVoiceBtn.disabled = true;
    setDiagnostic("Browser backup voice მიუწვდომელია. AI ხმა მაინც იმუშავებს, თუ /api/tts სწორადაა დაყენებული.");
    return;
  }

  voices = speech.getVoices();
  const selectedValue = voiceSelect.value;
  const georgianVoices = voices.filter((voice) => voice.lang?.toLowerCase().startsWith("ka"));
  const preferredVoices = [...georgianVoices, ...voices.filter((voice) => !voice.lang?.toLowerCase().startsWith("ka"))];

  voiceSelect.innerHTML = "";
  for (const voice of preferredVoices) {
    const option = document.createElement("option");
    option.value = voice.name;
    option.textContent = `${voice.name} (${voice.lang})`;
    voiceSelect.append(option);
  }

  if (!preferredVoices.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "ხმა ვერ მოიძებნა";
    voiceSelect.append(option);
  }

  if (selectedValue) {
    voiceSelect.value = selectedValue;
  }

  if (!voices.length) {
    setDiagnostic("AI ხმა არის მთავარი. Browser backup voice-ების სია ჯერ ცარიელია.");
    return;
  }

  if (georgianVoices.length) {
    setDiagnostic(`AI ხმა არის მთავარი. Browser backup-ად ქართული ხმაც ჩანს: ${georgianVoices[0].name}.`);
  } else {
    setDiagnostic(`AI ხმა არის მთავარი. Browser backup-ში ${voices.length} ხმაა, მაგრამ ქართული voice არ ჩანს.`);
  }
}

function selectedVoice() {
  return voices.find((voice) => voice.name === voiceSelect.value) || voices.find((voice) => voice.lang?.toLowerCase().startsWith("ka")) || null;
}

function cancelBrowserSpeech() {
  if (!speech) {
    return;
  }

  ignoreSpeechErrorsUntil = Date.now() + 900;
  speech.pause();
  speech.cancel();
  speech.resume();
  speech.cancel();
}

function stopReading() {
  readingSession += 1;
  clearTimeout(advanceTimer);
  activeTtsController?.abort();
  activeTtsController = null;
  cancelBrowserSpeech();
  if (activeAudio) {
    activeAudio.pause();
    activeAudio.src = "";
    activeAudio.load();
  }
  activeAudio = null;
  readingParts = [];
  readingPartIndex = 0;
  isReading = false;
  isVisualFallback = false;
  document.body.classList.remove("reading");
  playBtn.textContent = "▶";
  setupPlayBtn.textContent = "წაკითხვა";
}

async function speakWithBackend(text, onDone, options = {}) {
  const result = await fetch("/api/tts", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ text }),
    signal: options.signal,
  });

  if (options.signal?.aborted) {
    return;
  }

  if (!result.ok) {
    const message = await result.text();
    lastBackendError = message;
    throw new Error(message || "TTS backend failed");
  }

  const audioBlob = await result.blob();
  if (options.signal?.aborted) {
    return;
  }

  const audioUrl = URL.createObjectURL(audioBlob);
  const audio = new Audio(audioUrl);
  activeAudio = audio;

  const stopAudio = () => {
    audio.pause();
    audio.src = "";
    audio.load();
    URL.revokeObjectURL(audioUrl);
    if (activeAudio === audio) {
      activeAudio = null;
    }
  };

  options.signal?.addEventListener("abort", stopAudio, { once: true });

  audio.onplay = () => {
    if (!options.signal?.aborted) {
      onDone?.("start");
    }
  };
  audio.onended = () => {
    URL.revokeObjectURL(audioUrl);
    if (activeAudio === audio) {
      activeAudio = null;
    }
    if (!options.signal?.aborted) {
      onDone?.("end");
    }
  };
  audio.onerror = () => {
    URL.revokeObjectURL(audioUrl);
    if (activeAudio === audio) {
      activeAudio = null;
    }
    if (!options.signal?.aborted) {
      onDone?.("fail", "audio playback failed");
    }
  };

  if (options.signal?.aborted) {
    stopAudio();
    return;
  }

  await audio.play();
}

function speakText(text, onDone) {
  const session = readingSession;
  const controller = new AbortController();
  activeTtsController?.abort();
  activeTtsController = controller;
  const guardedDone = (state, reason) => {
    if (controller.signal.aborted || session !== readingSession) {
      return;
    }
    onDone?.(state, reason);
  };

  speakWithBackend(text, guardedDone, { signal: controller.signal }).catch((error) => {
    if (controller.signal.aborted || session !== readingSession) {
      return;
    }
    const message = error.message || lastBackendError || "unknown backend error";
    setDiagnostic(`AI backend ვერ ჩაირთო: ${message.slice(0, 180)}. ვცდი transliteration backup ხმას.`);
    if (/quota|billing|insufficient_quota/i.test(message)) {
      speakWithBrowser(transliterateGeorgian(text), guardedDone, { forceEnglish: true, signal: controller.signal });
      return;
    }

    speakWithBrowser(text, guardedDone, { signal: controller.signal });
  });
}

function speakWithBrowser(text, onDone, options = {}) {
  if (options.signal?.aborted) {
    return;
  }

  if (!speech) {
    onDone?.("fail", "ამ ბრაუზერს ხმით წაკითხვა არ შეუძლია");
    return;
  }

  const voice = selectedVoice();
  const voiceLang = voice?.lang || "";
  const shouldTransliterate = options.forceEnglish || (containsGeorgian(text) && !voiceLang.toLowerCase().startsWith("ka"));
  const spokenText = shouldTransliterate ? transliterateGeorgian(text) : text;
  const utterance = new SpeechSynthesisUtterance(spokenText);
  utterance.lang = options.forceEnglish ? "en-US" : voiceLang || "ka-GE";
  utterance.rate = 0.88;
  utterance.pitch = 1;

  if (voice) {
    utterance.voice = voice;
  }

  let started = false;
  let startedAt = 0;
  const startGuard = setTimeout(() => {
    if (!started && !options.signal?.aborted) {
      setStatus("ხმა არ დაიწყო. სცადე Chrome/Edge ან სხვა voice dropdown-იდან.");
      onDone?.("fail", "ხმა არ დაიწყო");
    }
  }, 1400);

  options.signal?.addEventListener("abort", () => {
    clearTimeout(startGuard);
    cancelBrowserSpeech();
  }, { once: true });

  utterance.onstart = () => {
    if (options.signal?.aborted) {
      cancelBrowserSpeech();
      return;
    }
    started = true;
    startedAt = performance.now();
    clearTimeout(startGuard);
    onDone?.("start");
  };

  utterance.onend = () => {
    clearTimeout(startGuard);
    if (options.signal?.aborted) {
      return;
    }
    if (startedAt && performance.now() - startedAt < 250) {
      onDone?.("fail", "ხმა მაშინვე დასრულდა");
      return;
    }
    onDone?.("end");
  };

  utterance.onerror = (event) => {
    clearTimeout(startGuard);
    const errorName = String(event.error || "").toLowerCase().trim();
    const isExpectedCancel = Date.now() < ignoreSpeechErrorsUntil || ["interrupted", "canceled", "cancelled"].includes(errorName);
    if (options.signal?.aborted || isExpectedCancel) {
      return;
    }
    stopReading();
    setStatus(`ხმის შეცდომა: ${event.error || "unknown"}. სცადე სხვა ბრაუზერი ან voice.`);
    onDone?.("fail", event.error || "unknown");
  };

  cancelBrowserSpeech();
  speech.speak(utterance);
}

function estimatedReadingTime(text) {
  const words = text.split(/\s+/).filter(Boolean).length;
  return Math.max(2200, Math.min(12000, words * 420));
}

function visualFallbackReadCurrentSlide(reason) {
  const session = readingSession;
  const script = getSlideScript(currentSlide);

  isReading = true;
  isVisualFallback = true;
  document.body.classList.add("reading");
  playBtn.textContent = "■";
  setupPlayBtn.textContent = "შეჩერება";
  setStatus(`ხმა არ ირთვება (${reason}). demo რეჟიმში სლაიდი ავტომატურად გადავა.`);

  clearTimeout(advanceTimer);
  advanceTimer = setTimeout(() => {
    if (session !== readingSession) {
      return;
    }
    finishSlideReading();
  }, estimatedReadingTime(script));
}

function splitForSpeech(text) {
  const sentences = text
    .replace(/\s+/g, " ")
    .split(/(?<=[.!?։…])\s+/)
    .map((part) => part.trim())
    .filter(Boolean);

  const parts = [];
  let current = "";

  for (const sentence of sentences.length ? sentences : [text]) {
    if (`${current} ${sentence}`.trim().length > 180 && current) {
      parts.push(current);
      current = sentence;
    } else {
      current = `${current} ${sentence}`.trim();
    }
  }

  if (current) {
    parts.push(current);
  }

  return parts;
}

function finishSlideReading() {
  const session = readingSession;
  advanceTimer = setTimeout(() => {
    if (session !== readingSession) {
      return;
    }

    if (currentSlide < slides.length - 1) {
      currentSlide += 1;
      renderSlide();
      readCurrentSlide();
      return;
    }

    stopReading();
    setStatus("პრეზენტაცია დასრულდა");
  }, 2000);
}

function speakNextPart() {
  const session = readingSession;
  const text = readingParts[readingPartIndex];

  if (isVisualFallback) {
    return;
  }

  if (!text) {
    finishSlideReading();
    return;
  }

  speakText(text, (state) => {
    if (session !== readingSession) {
      return;
    }

    if (state === "start") {
      isReading = true;
      document.body.classList.add("reading");
      playBtn.textContent = "■";
      setupPlayBtn.textContent = "შეჩერება";
      setStatus(`კითხულობს: ნაწილი ${readingPartIndex + 1} / ${readingParts.length}`);
    }

    if (state === "end") {
      readingPartIndex += 1;
      speakNextPart();
    }

    if (state === "fail") {
      visualFallbackReadCurrentSlide("TTS unavailable");
    }
  });
}

function readCurrentSlide() {
  const script = getSlideScript(currentSlide);

  if (!script) {
    setStatus("ამ სლაიდის ტექსტი ცარიელია");
    return;
  }

  clearTimeout(advanceTimer);
  activeTtsController?.abort();
  activeTtsController = null;
  cancelBrowserSpeech();
  readingSession += 1;
  readingParts = splitForSpeech(script);
  readingPartIndex = 0;
  isReading = true;
  isVisualFallback = false;
  document.body.classList.add("reading");
  playBtn.textContent = "■";
  setupPlayBtn.textContent = "შეჩერება";
  setStatus("ხმა მზადდება...");
  speakNextPart();
}

function goToSlide(nextIndex) {
  if (!slides.length) {
    return;
  }

  currentSlide = Math.max(0, Math.min(slides.length - 1, nextIndex));
  renderSlide();

  if (isReading) {
    readCurrentSlide();
  }
}

function handleAresCommand(rawCommand) {
  const command = rawCommand.toLowerCase().replace(/[.,!?]/g, " ").replace(/\s+/g, " ").trim();
  const wakeWords = ["ares", "aris", "erase", "heiress", "არესი", "არეს"];
  const wakeWord = wakeWords.find((word) => command.includes(word));

  if (!wakeWord) {
    setAresStatus(`Ares უსმენს... გავიგე: “${rawCommand}”`);
    return;
  }

  const commandWithoutWake = command.replace(wakeWord, "").trim();
  if (!commandWithoutWake) {
    stopReading();
    setAresStatus("Ares: გავჩერდი და გისმენ.");
    return;
  }

  if (/\b(next|forward)\b|შემდეგ|შემდეგი/.test(command)) {
    stopReading();
    goToSlide(currentSlide + 1);
    setAresStatus("Ares: შემდეგ სლაიდზე გადავედი.");
    return;
  }

  if (/\b(previous|prev|back)\b|წინა|უკან/.test(command)) {
    stopReading();
    goToSlide(currentSlide - 1);
    setAresStatus("Ares: წინა სლაიდზე დავბრუნდი.");
    return;
  }

  if (/\b(play|read|start|speak)\b|წაიკითხე|დაიწყე|გააგრძელე/.test(command)) {
    stopReading();
    setAresStatus("Ares: ვიწყებ წაკითხვას.");
    readCurrentSlide();
    return;
  }

  if (/\b(stop|pause|cancel|quiet|shut|halt|enough)\b|გაჩერდი|შეჩერდი|სდექ|გაჩერდეს|გააჩერე/.test(command)) {
    stopReading();
    setAresStatus("Ares: გავჩერდი და გისმენ.");
    return;
  }

  stopReading();
  setAresStatus("Ares: გავჩერდი. მითხარი play, next, previous ან stop.");
}

function startAres() {
  if (!SpeechRecognition) {
    setAresStatus("Ares ვერ ჩაირთო: ამ ბრაუზერს speech recognition არ აქვს.");
    return;
  }

  if (!aresRecognition) {
    aresRecognition = new SpeechRecognition();
    aresRecognition.lang = "en-US";
    aresRecognition.continuous = true;
    aresRecognition.interimResults = false;
    aresRecognition.maxAlternatives = 1;

    aresRecognition.addEventListener("result", (event) => {
      const result = event.results[event.results.length - 1];
      const transcript = result?.[0]?.transcript || "";
      if (transcript) {
        handleAresCommand(transcript);
      }
    });

    aresRecognition.addEventListener("end", () => {
      if (aresEnabled) {
        try {
          aresRecognition.start();
        } catch (error) {
          setAresStatus("Ares დროებით გაჩერდა. დააჭირე Ares-ს თავიდან.");
        }
      }
    });

    aresRecognition.addEventListener("error", (event) => {
      setAresStatus(`Ares შეცდომა: ${event.error}.`);
    });
  }

  aresEnabled = true;
  aresBtn.classList.add("listening");
  setAresStatus("Ares ჩართულია. თქვი: “Ares stop”, “Ares next”, “Ares play”.");

  try {
    aresRecognition.start();
  } catch (error) {
    setAresStatus("Ares უკვე უსმენს.");
  }
}

function stopAres() {
  aresEnabled = false;
  aresBtn.classList.remove("listening");
  aresRecognition?.stop();
  setAresStatus("Ares გამორთულია.");
}

pptInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) {
    return;
  }

  pptName.textContent = file.name;
  stopReading();
  setStatus("PowerPoint იტვირთება...");

  try {
    await loadPowerPoint(file);
  } catch (error) {
    setStatus(error.message || "PowerPoint ვერ ჩაიტვირთა");
  }
});

docInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) {
    return;
  }

  docName.textContent = file.name;
  stopReading();
  setStatus("Word ტექსტი იტვირთება...");

  try {
    await loadWord(file);
  } catch (error) {
    setStatus("Word ტექსტი ვერ ჩაიტვირთა");
  }
});

setupFileInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) {
    return;
  }

  if (file.name.toLowerCase().endsWith(".pptx")) {
    pptName.textContent = file.name;
    stopReading();
    setStatus("PowerPoint იტვირთება...");
    try {
      await loadPowerPoint(file);
    } catch (error) {
      setStatus(error.message || "PowerPoint ვერ ჩაიტვირთა");
    }
    return;
  }

  if (file.name.toLowerCase().endsWith(".docx")) {
    docName.textContent = file.name;
    stopReading();
    setStatus("Word ტექსტი იტვირთება...");
    try {
      await loadWord(file);
    } catch (error) {
      setStatus("Word ტექსტი ვერ ჩაიტვირთა");
    }
    return;
  }

  setStatus("მხოლოდ .pptx ან .docx ფაილი აირჩიე");
});

playBtn.addEventListener("click", () => {
  if (isReading) {
    stopReading();
    setStatus("წაკითხვა შეჩერდა");
    return;
  }

  readCurrentSlide();
});

setupPlayBtn.addEventListener("click", () => {
  playBtn.click();
});

prevBtn.addEventListener("click", () => goToSlide(currentSlide - 1));
nextBtn.addEventListener("click", () => goToSlide(currentSlide + 1));

aresBtn.addEventListener("click", () => {
  if (aresEnabled) {
    stopAres();
    return;
  }

  startAres();
});

testVoiceBtn.addEventListener("click", () => {
  speech?.cancel();
  setStatus("AI ხმის ტესტი იწყება...");
  speakText("გამარჯობა, ეს არის ქართული ხმის ტესტი.", (state, reason) => {
    if (state === "start") {
      setStatus("ხმა მუშაობს");
    }

    if (state === "end") {
      setStatus("ხმის ტესტი დასრულდა");
    }

    if (state === "fail") {
      setStatus(`ხმის ტესტი ვერ გავიდა: ${reason}. გადაამოწმე Vercel OPENAI_API_KEY.`);
    }
  });
});

speech?.addEventListener("voiceschanged", populateVoices);
populateVoices();
setTimeout(populateVoices, 600);
renderSlide();
