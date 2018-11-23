$('.carousel').carousel()
$(document).on('click', '.yamm .dropdown-menu', function(e) {
    e.stopPropagation();
})
/*Tooltip*/
$(function() {
    $('[data-toggle="tooltip"]').tooltip();
});




var upperMenuTot = $(".first-top-menu").outerHeight() + $(".last-top-menu").outerHeight();

console.log(upperMenuTot);

$(document).scroll(function (event) {
    var scroll = $(window).scrollTop();
    if (scroll >= upperMenuTot){
        // $(".content").css("margin-top", "130px");
        $('#top-link-block').removeClass('hidden');
    } else {
        // $(".content").css("margin-top", "5px");
        $('#top-link-block').addClass('hidden');

    };
});

var affixElement = '#block-main';

$(affixElement).affix({
    offset: {
        // Distance of between element and top page
        top: function() {
            return (this.top = $(affixElement).offset().top);
        },
        // when start #footer
        bottom: function() {
            return (this.bottom = $('#footer').outerHeight(true))
        }
    }
});
