
// from: https://stackoverflow.com/questions/6787383/how-to-add-remove-a-class-in-javascript
export function hasClass(element : HTMLElement, className : string) {
    return !!element.className.match(new RegExp('(\\s|^)' + className + '(\\s|$)'));
}

export function addClass(element: HTMLElement , className : string) {
    if (!hasClass(element, className)) {
        element.className += " " + className;
    }
}

export function removeClass(element : HTMLElement, className : string) {
    if ( hasClass(element, className) ) {
        var reg = new RegExp('(\\s|^)'+className+'(\\s|$)');
        element.className=element.className.replace(reg,' ');
    }
}

export function toggleClass(element: HTMLElement, className: string) {
    if (hasClass(element, className)) {
        var reg = new RegExp('(\\s|^)'+className+'(\\s|$)');
        element.className=element.className.replace(reg,' ');
    } else {
        element.className += " " + className;
    }
}