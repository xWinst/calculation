import { Report } from 'notiflix/build/notiflix-report-aio';
import { Notify } from 'notiflix/build/notiflix-notify-aio';

const styles = {
    width: '350px',
    svgSize: '100px',
    titleFontSize: '20px',
    buttonFontSize: '20px',
    borderRadius: '10px',
};

Report.init({
    plainText: false,
    titleMaxLength: 100,
});

class Message {
    warning(warning, text, buttonText = 'Ok') {
        return Report.warning(warning, text, buttonText, styles);
    }

    // error(error, text, buttonText = 'Ok') {
    //     return Report.failure(error, text, buttonText, styles);
    // }

    // success(text) {
    //     return Notify.success(text, { position: 'center-top' });
    // }
}
const message = new Message();

export default message;
