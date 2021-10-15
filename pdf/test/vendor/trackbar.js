window.initTrackBars = function()
{
    var trackbars = document.querySelectorAll('.trackbar');
    for (let i = 0; i < trackbars.length; i++)
    {
        let trackbar = trackbars[i];

        var button = trackbar.querySelector('.track-button');
        if (button && button instanceof HTMLElement)
        {
            let dragging = false;
            let clickOffset = 0;

            button.addEventListener('mousedown', function(event) {
                const parentWidth = trackbar.offsetWidth;
                const elementWidth = button.offsetWidth;

                clickOffset = (event.pageX - button.offsetLeft - 8) - (elementWidth / 2);
                dragging = true;
            });
            document.body.addEventListener('mouseup', function(event) {
                dragging = false;
            });

            document.body.addEventListener('mousemove', function(event) {
                if(dragging)
                {
                    const parentWidth = trackbar.offsetWidth;
                    const elementWidth = button.offsetWidth;

                    let currentPosition = (event.pageX - button.offsetLeft - 8) - (elementWidth / 2);
                    let elementPosition = parseInt(window.getComputedStyle(button).getPropertyValue('left'));

                    let nextPosition = elementPosition + currentPosition - clickOffset;

                    if(nextPosition < 0) nextPosition = 0;
                    if(nextPosition >= parentWidth - elementWidth) nextPosition = parentWidth - elementWidth;

                    if(elementPosition == nextPosition + 2) return false;
                    nextPosition = nextPosition + 'px';

                    button.style.left = nextPosition;
                    trackbar.setAttribute('complete', Math.floor(elementPosition * 100 / (parentWidth - elementWidth)));
                    if (trackSelector = trackbar.querySelector('.track-selector'))
                        trackSelector.style.width = Math.floor(elementPosition * parentWidth / (parentWidth - elementWidth)) + 'px';

                    if (trackPercentage = trackbar.querySelector('.track-percentage'))
                        trackPercentage.innerHTML = Math.floor(elementPosition * 100 / (parentWidth - elementWidth)) + '%';

                    if (trackbar.onChangedValue)
                        trackbar.onChangedValue(elementPosition  / (parentWidth - elementWidth));
                }
            });
            document.body.addEventListener('mouseleave', function(event) {dragging = false;});

            trackbar.setPosition = function(percent) {
                const parentWidth = trackbar.offsetWidth;
                const elementWidth = button.offsetWidth;

                let nextPosition = (percent * (parentWidth - elementWidth)) >> 0;
                let elementPosition = parseInt(window.getComputedStyle(button).getPropertyValue('left'));

                if(elementPosition == nextPosition + 2) return false;

                if(nextPosition < 0) nextPosition = 0;
                if(nextPosition >= parentWidth - elementWidth) nextPosition = parentWidth - elementWidth;

                if (trackSelector = trackbar.querySelector('.track-selector'))
                    trackSelector.style.width = Math.floor(elementPosition * parentWidth / (parentWidth - elementWidth)) + 'px';

                if (trackPercentage = trackbar.querySelector('.track-percentage'))
                    trackPercentage.innerHTML = Math.floor(elementPosition * 100 / (parentWidth - elementWidth)) + '%';

                button.style.left = nextPosition + "px";                
            };
        }
    }
}
