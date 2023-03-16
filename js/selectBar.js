function prepColumns() {
  let c = document.querySelectorAll(".XLInner")

  for (let i = 0; i < c.length; i++) {
    c[i].addEventListener('click', function() {
      let row = this.className.split('XLInner ')[1];
      if (row.includes("blank")) row = row.split("blank ")[1];

      while (document.getElementsByClassName("selected").length > 0) {
        document.getElementsByClassName("selected")[0].classList.toggle("selected");
      }

      for (let j = 0; j < document.getElementsByClassName(row).length; j++) {
        document.getElementsByClassName(row)[j].classList.toggle("selected");
      }
    })
  }
}