const uploadArea = document.getElementById("uploadArea");
const fileInput = document.getElementById("excelFile");
const fileInfo = document.getElementById("fileInfo");
const fileName = document.getElementById("fileName");
const form = document.getElementById("calculationForm");
const loading = document.getElementById("loading");
const result = document.getElementById("result");
const error = document.getElementById("error");
const submitBtn = document.getElementById("submitBtn");
const downloadBtn = document.getElementById("downloadBtn");

uploadArea.addEventListener("dragover", (event) => {
    event.preventDefault();
    uploadArea.classList.add("dragover");
});

uploadArea.addEventListener("dragleave", () => {
    uploadArea.classList.remove("dragover");
});

uploadArea.addEventListener("drop", (event) => {
    event.preventDefault();
    uploadArea.classList.remove("dragover");
    if (event.dataTransfer.files.length > 0) {
        fileInput.files = event.dataTransfer.files;
        showFileName(event.dataTransfer.files[0].name);
    }
});

fileInput.addEventListener("change", (event) => {
    if (event.target.files.length > 0) {
        showFileName(event.target.files[0].name);
    }
});

function showFileName(name) {
    fileName.textContent = name;
    fileInfo.classList.add("show");
}

form.addEventListener("submit", async (event) => {
    event.preventDefault();
    result.classList.remove("show");
    error.classList.remove("show");
    loading.classList.add("show");
    submitBtn.disabled = true;

    try {
        const response = await fetch("/hesapla", {
            method: "POST",
            body: new FormData(form),
        });
        const data = await response.json();
        loading.classList.remove("show");

        if (response.ok && data.success) {
            document.getElementById("sabitKiymetSayisi").textContent = `${data.sabit_kiymet_sayisi} adet`;
            document.getElementById("ydFisSayisi").textContent = `${data.yd_fis_sayisi} adet`;
            document.getElementById("amortismanFisSayisi").textContent = `${data.amortisman_fis_sayisi} adet`;
            downloadBtn.onclick = () => {
                window.location.href = data.download_url;
            };
            result.classList.add("show");
        } else {
            error.textContent = data.error || "Bir hata oluştu.";
            error.classList.add("show");
        }
    } catch (err) {
        loading.classList.remove("show");
        error.textContent = `Bağlantı hatası: ${err.message}`;
        error.classList.add("show");
    } finally {
        submitBtn.disabled = false;
    }
});
