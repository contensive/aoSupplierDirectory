jQuery(document).ready(function() {
    $('#js-members-carousel').slick({
        infinite: true,
        slidesToShow: 5,
        slidesToScroll: 1,
        autoplay: true,
        autoplaySpeed: 3500,
        arrows: true,
        prevArrow: '<img src="img/carou-arr-L.png" class="slick-prev" alt="Previous" />',
        nextArrow: '<img src="img/carou-arr-R.png" class="slick-next" alt="Next" />',
        responsive: [
            {
                breakpoint: 1024,
                settings: {
                    slidesToShow: 4,
                    slidesToScroll: 4
                }
            },
            {
                breakpoint: 768,
                settings: {
                    slidesToShow: 3,
                    slidesToScroll: 3
                }
            },
            {
                breakpoint: 480,
                settings: {
                    slidesToShow: 1,
                    slidesToScroll: 1
                }
            }
        ]
    });
});