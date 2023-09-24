const body = document.querySelector("body");
const modal = document.querySelector(".modal");
const modalButton = document.querySelector(".modal-button");
const closeButton = document.querySelector(".close-button");
const inputButton = document.querySelector(".input-button");
const scrollDown = document.querySelector(".scroll-down");
const errMsg = document.querySelector(".errMsg");
let isOpened = false;

const openModal = () => {
  modal.classList.add("is-open");
  body.style.overflow = "hidden";
};

const closeModal = () => {
  modal.classList.remove("is-open");
  body.style.overflow = "initial";
};

window.addEventListener("scroll", () => {
  if (window.scrollY > window.innerHeight / 3 && !isOpened) {
    isOpened = true;
    scrollDown.style.display = "none";
    openModal();
  }
});

const checkErr = () => {
  if (errMsg) {
    openModal();
  }
};


const inputModal = () => {
  loginForm = document.getElementById("loginForm");
  loginForm.submit();
};

modalButton.addEventListener("click", openModal);
closeButton.addEventListener("click", closeModal);
inputButton.addEventListener("click", inputModal);


document.onkeydown = evt => {
  evt = evt || window.event;
  evt.keyCode === 27 ? closeModal() : false;
};
