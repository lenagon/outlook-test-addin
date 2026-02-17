Office.onReady(() => {
  loadSettings();

  document.getElementById("saveBtn").addEventListener("click", saveSettings);
});

function loadSettings() {
  const settings = Office.context.roamingSettings;

  document.getElementById("serverUrl").value =
    settings.get("serverUrl") || "";

  document.getElementById("username").value =
    settings.get("username") || "";

  document.getElementById("password").value =
    settings.get("password") || "";
}

function saveSettings() {
  const settings = Office.context.roamingSettings;

  settings.set("serverUrl",
    document.getElementById("serverUrl").value);

  settings.set("username",
    document.getElementById("username").value);

  settings.set("password",
    document.getElementById("password").value);

  settings.saveAsync(result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      document.getElementById("status").innerText =
        "Settings saved successfully.";
    } else {
      document.getElementById("status").innerText =
        "Error saving settings.";
    }
  });
}
