let flashMessageTimer: number | undefined;

export default function showFlashMessage(message: string) {
  const flashMessage = document.getElementById("flash-message");

  if (!flashMessage) {
    console.log(message);
    return;
  }

  flashMessage.textContent = message;
  flashMessage.classList.add("is-visible");

  if (flashMessageTimer !== undefined) {
    window.clearTimeout(flashMessageTimer);
  }

  flashMessageTimer = window.setTimeout(() => {
    flashMessage.textContent = "";
    flashMessage.classList.remove("is-visible");
    flashMessageTimer = undefined;
  }, 4000);
}