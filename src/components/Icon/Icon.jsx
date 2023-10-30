import icons from "icons/icons.svg";

const Icon = ({ icon, cn, onClick, w, h = w, s }) => {
    return (
        <svg className={cn} onClick={onClick} width={w} height={h} style={s}>
            <use href={`${icons}#${icon}`} />
        </svg>
    );
};

export default Icon;
