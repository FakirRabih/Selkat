$(document).ready(function () {
    $('#start-logging-btn').click(function () {
        $.ajax({
            url: 'http://localhost:3000/start-logging',
            method: 'POST',
            success: function () {
                console.log('Logging started');
            },
            error: function (error) {
                console.error('Error:', error);
            }
        });
    });
});