/* ============================================================
   Conversion Heaven – Application Logic
   ============================================================ */

(function () {
    "use strict";

    /* ---------- DOM References ---------- */
    const themeToggle = document.getElementById("themeToggle");
    const themeIcon = document.getElementById("themeIcon");
    const loadingOverlay = document.getElementById("loadingOverlay");
    const thankYouModal = document.getElementById("thankYouModal");
    const modalCloseBtn = document.getElementById("modalCloseBtn");
    const particleCanvas = document.getElementById("particleCanvas");
    const ctx = particleCanvas.getContext("2d");

    /* ---------- Theme Toggle ---------- */
    function applyTheme(theme) {
        document.documentElement.setAttribute("data-theme", theme);
        themeIcon.textContent = theme === "dark" ? "🌙" : "☀️";
        localStorage.setItem("ch-theme", theme);
    }

    themeToggle.addEventListener("click", function () {
        var current = document.documentElement.getAttribute("data-theme");
        applyTheme(current === "dark" ? "light" : "dark");
    });

    var stored = localStorage.getItem("ch-theme");
    if (stored) applyTheme(stored);

    /* ---------- Particle System ---------- */
    var particles = [];
    var EMOJIS = ["🎉", "✨", "🥳", "🎊", "⭐", "💜", "🔥", "🚀"];

    function resizeCanvas() {
        particleCanvas.width = window.innerWidth;
        particleCanvas.height = window.innerHeight;
    }
    window.addEventListener("resize", resizeCanvas);
    resizeCanvas();

    function spawnParticles(x, y) {
        for (var i = 0; i < 30; i++) {
            particles.push({
                x: x,
                y: y,
                vx: (Math.random() - 0.5) * 12,
                vy: (Math.random() - 0.5) * 12 - 4,
                life: 1,
                emoji: EMOJIS[Math.floor(Math.random() * EMOJIS.length)],
                size: 16 + Math.random() * 18,
                rotation: Math.random() * 360,
                rotationSpeed: (Math.random() - 0.5) * 10
            });
        }
    }

    function animateParticles() {
        ctx.clearRect(0, 0, particleCanvas.width, particleCanvas.height);
        for (var i = particles.length - 1; i >= 0; i--) {
            var p = particles[i];
            p.x += p.vx;
            p.y += p.vy;
            p.vy += 0.25;
            p.life -= 0.018;
            p.rotation += p.rotationSpeed;
            if (p.life <= 0) {
                particles.splice(i, 1);
                continue;
            }
            ctx.save();
            ctx.globalAlpha = p.life;
            ctx.translate(p.x, p.y);
            ctx.rotate((p.rotation * Math.PI) / 180);
            ctx.font = p.size + "px serif";
            ctx.textAlign = "center";
            ctx.textBaseline = "middle";
            ctx.fillText(p.emoji, 0, 0);
            ctx.restore();
        }
        requestAnimationFrame(animateParticles);
    }
    animateParticles();

    /* ---------- Loading helpers ---------- */
    function showLoading() {
        loadingOverlay.classList.add("active");
    }
    function hideLoading() {
        loadingOverlay.classList.remove("active");
    }

    /* ---------- Thank-you modal ---------- */
    function showThankYou() {
        thankYouModal.classList.add("active");
    }
    modalCloseBtn.addEventListener("click", function () {
        thankYouModal.classList.remove("active");
    });

    /* ---------- Download helper ---------- */
    function triggerDownload(blob, filename, btnEvent) {
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        /* Revoke immediately – "delete from database after conversion" */
        setTimeout(function () { URL.revokeObjectURL(url); }, 500);

        /* Emoji particles at click position */
        if (btnEvent) {
            spawnParticles(btnEvent.clientX, btnEvent.clientY);
        } else {
            spawnParticles(window.innerWidth / 2, window.innerHeight / 2);
        }

        /* Show thank-you modal after a brief pause */
        setTimeout(showThankYou, 600);
    }

    /* ---------- Helper: read file as ArrayBuffer ---------- */
    function readFileAsArrayBuffer(file) {
        return new Promise(function (resolve, reject) {
            var reader = new FileReader();
            reader.onload = function () { resolve(reader.result); };
            reader.onerror = function () { reject(reader.error); };
            reader.readAsArrayBuffer(file);
        });
    }

    /* ---------- Helper: read file as Text ---------- */
    function readFileAsText(file) {
        return new Promise(function (resolve, reject) {
            var reader = new FileReader();
            reader.onload = function () { resolve(reader.result); };
            reader.onerror = function () { reject(reader.error); };
            reader.readAsText(file);
        });
    }

    /* ---------- Helper: read file as Data URL ---------- */
    function readFileAsDataURL(file) {
        return new Promise(function (resolve, reject) {
            var reader = new FileReader();
            reader.onload = function () { resolve(reader.result); };
            reader.onerror = function () { reject(reader.error); };
            reader.readAsDataURL(file);
        });
    }

    /* ---------- Library availability check ---------- */
    function checkLibraries(names) {
        var missing = [];
        var libs = {
            "jspdf": function () { return window.jspdf && window.jspdf.jsPDF; },
            "mammoth": function () { return typeof mammoth !== "undefined"; },
            "htmlDocx": function () { return window.htmlDocx && window.htmlDocx.asBlob; },
            "XLSX": function () { return typeof XLSX !== "undefined"; },
            "pdfjsLib": function () { return typeof pdfjsLib !== "undefined"; }
        };
        for (var i = 0; i < names.length; i++) {
            if (libs[names[i]] && !libs[names[i]]()) {
                missing.push(names[i]);
            }
        }
        if (missing.length > 0) {
            throw new Error(
                "Required libraries failed to load: " + missing.join(", ") +
                ". Please check your internet connection and reload the page."
            );
        }
    }

    /* ---------- Conversion: Word → PDF ---------- */
    async function convertWordToPdf(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["mammoth", "jspdf"]);
            var arrayBuffer = await readFileAsArrayBuffer(file);
            var result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
            var html = result.value;

            var jsPDF = window.jspdf.jsPDF;
            var doc = new jsPDF({ unit: "pt", format: "a4" });

            /* Strip HTML to plain text for jsPDF */
            var tempDiv = document.createElement("div");
            tempDiv.innerHTML = html;
            var text = tempDiv.textContent || tempDiv.innerText || "";

            var lines = doc.splitTextToSize(text, 500);
            var y = 40;
            var pageHeight = doc.internal.pageSize.getHeight();
            for (var i = 0; i < lines.length; i++) {
                if (y > pageHeight - 40) {
                    doc.addPage();
                    y = 40;
                }
                doc.text(lines[i], 40, y);
                y += 16;
            }

            var blob = doc.output("blob");
            triggerDownload(blob, file.name.replace(/\.docx?$/i, "") + ".pdf", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: PDF → Word ---------- */
    async function convertPdfToWord(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["pdfjsLib", "htmlDocx"]);
            var arrayBuffer = await readFileAsArrayBuffer(file);
            var pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            var fullText = "";
            for (var i = 1; i <= pdf.numPages; i++) {
                var page = await pdf.getPage(i);
                var content = await page.getTextContent();
                var pageText = content.items.map(function (item) { return item.str; }).join(" ");
                fullText += "<p>" + pageText + "</p>";
            }
            var converted = window.htmlDocx.asBlob("<html><body>" + fullText + "</body></html>");
            triggerDownload(converted, file.name.replace(/\.pdf$/i, "") + ".docx", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: PPT → PDF ---------- */
    async function convertPptToPdf(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["jspdf", "XLSX"]);
            var arrayBuffer = await readFileAsArrayBuffer(file);
            var jsPDF = window.jspdf.jsPDF;
            var doc = new jsPDF({ unit: "pt", format: "a4" });

            /* Read PPT XML content using JSZip embedded in xlsx */
            var workbook;
            try {
                workbook = XLSX.read(arrayBuffer, { type: "array" });
            } catch (_e) {
                /* If XLSX can't read it, just extract raw text */
                workbook = null;
            }

            var text = "";
            if (workbook && workbook.SheetNames && workbook.SheetNames.length > 0) {
                workbook.SheetNames.forEach(function (name) {
                    var sheet = workbook.Sheets[name];
                    text += XLSX.utils.sheet_to_txt(sheet) + "\n\n";
                });
            } else {
                /* Fallback: treat as binary, extract visible text */
                var bytes = new Uint8Array(arrayBuffer);
                var raw = "";
                for (var i = 0; i < bytes.length; i++) {
                    var ch = bytes[i];
                    if (ch >= 32 && ch < 127) raw += String.fromCharCode(ch);
                    else if (ch === 10 || ch === 13) raw += "\n";
                }
                text = raw.replace(/[^\x20-\x7E\n]/g, "").replace(/\n{3,}/g, "\n\n");
            }

            if (!text.trim()) {
                text = "(Slide content was extracted from " + file.name + ")";
            }

            var lines = doc.splitTextToSize(text, 500);
            var y = 40;
            var pageHeight = doc.internal.pageSize.getHeight();
            for (var j = 0; j < lines.length; j++) {
                if (y > pageHeight - 40) {
                    doc.addPage();
                    y = 40;
                }
                doc.text(lines[j], 40, y);
                y += 16;
            }

            var blob = doc.output("blob");
            triggerDownload(blob, file.name.replace(/\.pptx?$/i, "") + ".pdf", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: Excel → PDF ---------- */
    async function convertExcelToPdf(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["jspdf", "XLSX"]);
            var arrayBuffer = await readFileAsArrayBuffer(file);
            var workbook = XLSX.read(arrayBuffer, { type: "array" });

            var jsPDF = window.jspdf.jsPDF;
            var doc = new jsPDF({ unit: "pt", format: "a4" });

            var y = 40;
            var pageHeight = doc.internal.pageSize.getHeight();

            workbook.SheetNames.forEach(function (name, idx) {
                if (idx > 0) {
                    doc.addPage();
                    y = 40;
                }
                doc.setFontSize(14);
                doc.text("Sheet: " + name, 40, y);
                y += 24;
                doc.setFontSize(10);

                var sheet = workbook.Sheets[name];
                var text = XLSX.utils.sheet_to_csv(sheet);
                var lines = doc.splitTextToSize(text, 500);
                for (var i = 0; i < lines.length; i++) {
                    if (y > pageHeight - 40) {
                        doc.addPage();
                        y = 40;
                    }
                    doc.text(lines[i], 40, y);
                    y += 14;
                }
            });

            var blob = doc.output("blob");
            triggerDownload(blob, file.name.replace(/\.xlsx?$/i, "") + ".pdf", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: Image → PDF ---------- */
    async function convertImageToPdf(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["jspdf"]);
            var dataUrl = await readFileAsDataURL(file);
            var jsPDF = window.jspdf.jsPDF;
            var doc = new jsPDF({ unit: "pt", format: "a4" });

            var img = new Image();
            await new Promise(function (resolve, reject) {
                img.onload = resolve;
                img.onerror = reject;
                img.src = dataUrl;
            });

            var pageW = doc.internal.pageSize.getWidth() - 80;
            var pageH = doc.internal.pageSize.getHeight() - 80;
            var ratio = Math.min(pageW / img.width, pageH / img.height, 1);
            var w = img.width * ratio;
            var h = img.height * ratio;
            doc.addImage(dataUrl, "JPEG", 40, 40, w, h);

            var blob = doc.output("blob");
            triggerDownload(blob, file.name.replace(/\.[^.]+$/, "") + ".pdf", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: PDF → Image ---------- */
    async function convertPdfToImage(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["pdfjsLib"]);
            var arrayBuffer = await readFileAsArrayBuffer(file);
            var pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            var page = await pdf.getPage(1);
            var scale = 2;
            var viewport = page.getViewport({ scale: scale });
            var canvas = document.createElement("canvas");
            canvas.width = viewport.width;
            canvas.height = viewport.height;
            var context = canvas.getContext("2d");
            await page.render({ canvasContext: context, viewport: viewport }).promise;

            await new Promise(function (resolve) {
                canvas.toBlob(function (blob) {
                    triggerDownload(blob, file.name.replace(/\.pdf$/i, "") + ".png", btnEvent);
                    resolve();
                }, "image/png");
            });
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: TXT → PDF ---------- */
    async function convertTxtToPdf(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["jspdf"]);
            var text = await readFileAsText(file);
            var jsPDF = window.jspdf.jsPDF;
            var doc = new jsPDF({ unit: "pt", format: "a4" });
            doc.setFontSize(11);
            var lines = doc.splitTextToSize(text, 500);
            var y = 40;
            var pageHeight = doc.internal.pageSize.getHeight();
            for (var i = 0; i < lines.length; i++) {
                if (y > pageHeight - 40) {
                    doc.addPage();
                    y = 40;
                }
                doc.text(lines[i], 40, y);
                y += 15;
            }
            var blob = doc.output("blob");
            triggerDownload(blob, file.name.replace(/\.txt$/i, "") + ".pdf", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: CSV → Excel ---------- */
    async function convertCsvToExcel(file, btnEvent) {
        showLoading();
        try {
            checkLibraries(["XLSX"]);
            var text = await readFileAsText(file);
            var workbook = XLSX.read(text, { type: "string" });
            var wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
            var blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
            triggerDownload(blob, file.name.replace(/\.csv$/i, "") + ".xlsx", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: Video → Audio ---------- */
    async function convertVideoToAudio(file, btnEvent) {
        showLoading();
        try {
            var url = URL.createObjectURL(file);
            var video = document.createElement("video");
            video.src = url;
            video.muted = true;

            await new Promise(function (resolve, reject) {
                video.onloadedmetadata = resolve;
                video.onerror = reject;
            });

            var audioCtx = new (window.AudioContext || window.webkitAudioContext)();
            var source = audioCtx.createMediaElementSource(video);
            var dest = audioCtx.createMediaStreamDestination();
            source.connect(dest);

            var recorder = new MediaRecorder(dest.stream, { mimeType: "audio/webm" });
            var chunks = [];
            recorder.ondataavailable = function (e) { if (e.data.size > 0) chunks.push(e.data); };

            var done = new Promise(function (resolve) { recorder.onstop = resolve; });
            recorder.start();
            video.currentTime = 0;
            video.muted = false;
            await video.play();

            /* Wait for video to finish */
            await new Promise(function (resolve) { video.onended = resolve; });
            recorder.stop();
            await done;

            audioCtx.close();
            URL.revokeObjectURL(url);

            var blob = new Blob(chunks, { type: "audio/webm" });
            triggerDownload(blob, file.name.replace(/\.[^.]+$/, "") + ".webm", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion: Audio → Video ---------- */
    async function convertAudioToVideo(file, btnEvent) {
        showLoading();
        try {
            var url = URL.createObjectURL(file);
            var audio = document.createElement("audio");
            audio.src = url;

            await new Promise(function (resolve, reject) {
                audio.onloadedmetadata = resolve;
                audio.onerror = reject;
            });

            /* Create a simple canvas with a waveform-like visual */
            var canvas = document.createElement("canvas");
            canvas.width = 1280;
            canvas.height = 720;
            var canvasCtx = canvas.getContext("2d");

            var audioCtx = new (window.AudioContext || window.webkitAudioContext)();
            var source = audioCtx.createMediaElementSource(audio);
            var analyser = audioCtx.createAnalyser();
            analyser.fftSize = 256;
            source.connect(analyser);
            analyser.connect(audioCtx.destination);

            var videoStream = canvas.captureStream(30);
            /* Merge audio */
            var audioDest = audioCtx.createMediaStreamDestination();
            source.connect(audioDest);
            audioDest.stream.getAudioTracks().forEach(function (t) { videoStream.addTrack(t); });

            var recorder = new MediaRecorder(videoStream, { mimeType: "video/webm" });
            var chunks = [];
            recorder.ondataavailable = function (e) { if (e.data.size > 0) chunks.push(e.data); };
            var done = new Promise(function (resolve) { recorder.onstop = resolve; });

            recorder.start();
            await audio.play();

            /* Draw visualizer */
            var bufferLength = analyser.frequencyBinCount;
            var dataArray = new Uint8Array(bufferLength);
            var animId;
            function draw() {
                animId = requestAnimationFrame(draw);
                analyser.getByteFrequencyData(dataArray);
                canvasCtx.fillStyle = "#1a0030";
                canvasCtx.fillRect(0, 0, canvas.width, canvas.height);
                var barWidth = (canvas.width / bufferLength) * 2.5;
                var x = 0;
                for (var i = 0; i < bufferLength; i++) {
                    var barHeight = dataArray[i] * 2;
                    canvasCtx.fillStyle = "hsl(" + (i * 3) + ", 80%, 60%)";
                    canvasCtx.fillRect(x, canvas.height - barHeight, barWidth, barHeight);
                    x += barWidth + 1;
                }
            }
            draw();

            await new Promise(function (resolve) { audio.onended = resolve; });
            cancelAnimationFrame(animId);
            recorder.stop();
            await done;

            audioCtx.close();
            URL.revokeObjectURL(url);

            var blob = new Blob(chunks, { type: "video/webm" });
            triggerDownload(blob, file.name.replace(/\.[^.]+$/, "") + ".webm", btnEvent);
        } catch (err) {
            alert("Conversion failed: " + err.message);
        } finally {
            hideLoading();
        }
    }

    /* ---------- Conversion dispatcher ---------- */
    var CONVERTERS = {
        "word-to-pdf": convertWordToPdf,
        "pdf-to-word": convertPdfToWord,
        "ppt-to-pdf": convertPptToPdf,
        "excel-to-pdf": convertExcelToPdf,
        "image-to-pdf": convertImageToPdf,
        "pdf-to-image": convertPdfToImage,
        "txt-to-pdf": convertTxtToPdf,
        "csv-to-excel": convertCsvToExcel,
        "video-to-audio": convertVideoToAudio,
        "audio-to-video": convertAudioToVideo
    };

    /* ---------- Wire up cards ---------- */
    document.querySelectorAll(".card").forEach(function (card) {
        var type = card.getAttribute("data-conversion");
        var input = card.querySelector("input[type=file]");
        var fileNameSpan = card.querySelector(".file-name");
        var convertBtn = card.querySelector(".convert-btn");
        var dropZone = card.querySelector(".drop-zone");
        var selectedFile = null;

        function selectFile(file) {
            selectedFile = file;
            fileNameSpan.textContent = file.name;
            convertBtn.disabled = false;
        }

        input.addEventListener("change", function () {
            if (input.files && input.files.length > 0) {
                selectFile(input.files[0]);
            }
        });

        /* ---------- Drag & Drop ---------- */
        if (dropZone) {
            var dragCounter = 0;
            var acceptAttr = input.getAttribute("accept") || "";

            function isFileAccepted(file) {
                if (!acceptAttr) return true;
                var accepts = acceptAttr.split(",").map(function (s) { return s.trim().toLowerCase(); });
                var fileName = file.name.toLowerCase();
                var fileType = (file.type || "").toLowerCase();
                return accepts.some(function (a) {
                    if (a.startsWith(".")) {
                        return fileName.endsWith(a);
                    }
                    if (a.endsWith("/*")) {
                        return fileType.startsWith(a.replace("/*", "/"));
                    }
                    return fileType === a;
                });
            }

            dropZone.addEventListener("dragover", function (e) {
                e.preventDefault();
                e.stopPropagation();
            });

            dropZone.addEventListener("dragenter", function (e) {
                e.preventDefault();
                e.stopPropagation();
                dragCounter++;
                dropZone.classList.add("drag-over");
            });

            dropZone.addEventListener("dragleave", function (e) {
                e.preventDefault();
                e.stopPropagation();
                dragCounter--;
                if (dragCounter <= 0) {
                    dragCounter = 0;
                    dropZone.classList.remove("drag-over");
                }
            });

            dropZone.addEventListener("drop", function (e) {
                e.preventDefault();
                e.stopPropagation();
                dragCounter = 0;
                dropZone.classList.remove("drag-over");
                if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                    var droppedFile = e.dataTransfer.files[0];
                    if (isFileAccepted(droppedFile)) {
                        selectFile(droppedFile);
                    } else {
                        alert("Invalid file type. Please drop a supported file (" + acceptAttr + ").");
                    }
                }
            });
        }

        convertBtn.addEventListener("click", function (e) {
            if (!selectedFile) return;
            var converter = CONVERTERS[type];
            if (converter) {
                converter(selectedFile, e);
                /* Clear file reference after conversion starts ("delete from database") */
                setTimeout(function () {
                    selectedFile = null;
                    input.value = "";
                    fileNameSpan.textContent = "";
                    convertBtn.disabled = true;
                }, 200);
            }
        });
    });

})();
