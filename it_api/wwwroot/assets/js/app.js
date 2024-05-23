/**
 * Theme: Metrica - Responsive Bootstrap 4 Admin Dashboard
 * Author: Mannatthemes
 * Module/App: Main Js
 */

function initMetisMenu() {
  //metis menu
  $(".metismenu").metisMenu();
}

function initLeftMenuCollapse() {
  // Left menu collapse
  $(document).on("click", ".button-menu-mobile", function (event) {
    event.preventDefault();
    $("body").toggleClass("enlarge-menu");
  });
}

function initEnlarge() {
  if ($(window).width() < 1025) {
    $("body").addClass("enlarge-menu");
  } else {
    if ($("body").data("keep-enlarged") != true)
      $("body").removeClass("enlarge-menu");
  }
}
function initActiveMenu() {
  // === following js will activate the menu in left side bar based on url ====
  $(".left-sidenav a").each(function () {
    var pageUrl = window.location.href.split(/[?#]/)[0];
    if (this.href == pageUrl) {
      $(this).addClass("active");
      $(this).parent().addClass("active"); // add active to li of the current link
      $(this).parent().addClass("mm-active");
      $(this).parent().parent().addClass("in");
      $(this).parent().parent().addClass("mm-show");
      $(this).parent().parent().parent().addClass("mm-active");
      $(this).parent().parent().prev().addClass("active"); // add active class to an anchor
      $(this).parent().parent().parent().addClass("active");
      $(this).parent().parent().parent().parent().addClass("mm-show"); // add active to li of the current link
      $(this)
        .parent()
        .parent()
        .parent()
        .parent()
        .parent()
        .addClass("mm-active");
    }
  });
}

function init() {
  //   initMetisMenu();
  //   initActiveMenu();
//   initLeftMenuCollapse();
//   initEnlarge();
  ////Confirm
//   $(document)
//     .off("click", "[data-type='confirm']")
//     .on("click", "[data-type='confirm']", function (e) {
//       e.preventDefault();
//       var title = $(this).attr("title");
//       var href = $(this).attr("href");
//       if (confirm(title) == true) {
//         if (href) location.href = href;
//       }
//       return false;
//     });
  // if (!$(".wait_loading").length) {
  //     $(".preloader").fadeOut();

  //     if ($(".chosen").length) {
  //         $(".chosen").chosen({
  //             search_contains: true
  //         });
  //     }
  //     if ($(".dropify").length) {
  //         $('.dropify').dropify();
  //     }
  //     if ($(".money").length) {
  //         $(".money").inputmask("numeric", {
  //             radixPoint: ".",
  //             groupSeparator: ",",
  //             autoGroup: true,
  //             suffix: ' VND', //No Space, this will truncate the first character
  //             rightAlign: false,
  //             oncleared: function () {
  //                 self.Value('');
  //             }
  //         });
  //     }
  // }

  //Swal.fire(
  //	'Thông báo bảo trì!',
  //	'Hệ thống sẽ bảo trì lúc 7h30 sáng ngày 15/11/2022 và dự kiến sẽ mở lại lúc 11h sáng ngày 15/11/2022.',
  //	'warning'
  //)
}

init();
