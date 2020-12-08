/**
 * @file
 */


(function ($) {
    Drupal.behaviors.owl = {
        attach: function (context, settings) {
            $('.owl-slider-wrapper', context).each(function () {
                var $this = $(this);
		        tmp = $this.attr('data-settings');
                var settings = $.parseJSON(tmp);

// https://stackoverflow.com/questions/38409040/how-to-add-captions-for-images-in-owl-carousel

                settings['afterMove'] = function(elem) {
                    var current = this.currentItem;
                    var currentImg = elem.find('.owl-item').eq(current).find('img');
                    $('figure')
                        .find('img')
                        .attr({
                            'src': currentImg.attr('src'),
                            'alt': currentImg.attr('alt'),
                            'title': currentImg.attr('title')
                        });
                    $('.owl_title').text(currentImg.attr('title'));
                    $('.owl_content').text(currentImg.attr('alt'));
                };

                settings['afterInit'] = function(elem) {
                    var current = this.currentItem;
                    var currentImg = elem.find('.owl-item').eq(current).find('img');
                    $('figure')
                        .find('img')
                        .attr({
                            'src': currentImg.attr('src'),
                            'alt': currentImg.attr('alt'),
                            'title': currentImg.attr('title')
                        });
                    $('.owl_title').text(currentImg.attr('title'));
                    $('.owl_content').text(currentImg.attr('alt'));
                };



                console.log(settings);
                $this.owlCarousel(settings);
            });
        }
    };

})(jQuery);    


/*




var jQueryScript = document.createElement('script'); jQueryScript.setAttribute('src','https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js'); document.head.appendChild(jQueryScript);


*/

/*
    loop: true,
    items: 1,
    navigation: true,
    pagination: true,
    lazyLoad: true,
    singleItem: true,
    afterMove: function(elem) {
        var current = this.currentItem;
        var currentImg = elem.find('.owl-item').eq(current).find('img');

        $('figure')
            .find('img')
            .attr({
                'src': currentImg.attr('src'),
                'alt': currentImg.attr('alt'),
                'title': currentImg.attr('title')
            });
        $('.owl-pagination').text(currentImg.attr('title'));
    }
});
*/





