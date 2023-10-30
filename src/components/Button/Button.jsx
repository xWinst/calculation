import { Icon } from 'components';
import s from './Button.module.css';

const Button = ({ cn, icon, text, click, type = 'button', st = {}, dis = false }) => {
    return (
        <button
            className={s.button + ' ' + cn}
            type={type}
            onClick={click}
            style={st}
            disabled={dis}
        >
            {icon && <Icon w="18" h="18" icon={icon} />}
            {text}
        </button>
    );
};

export default Button;
