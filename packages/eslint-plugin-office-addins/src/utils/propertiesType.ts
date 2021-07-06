const metadata = require('./metadata.json');

enum PropertyType {
    navigational,
    scalar,
    notProperty,
};

const navigationProperties: Set<string> = new Set<string> ();
const scalarProperties: Set<string> = new Set<string> ();


metadata.classes.forEach((classe: any) => {
    classe.properties.forEach((property: any) => {
        if (property.navigational === true) {
            navigationProperties.add(property.name);
        } else if (property.navigational === false) {
            scalarProperties.add(property.name);
        }
    });
});

export function getPropertyType(propertyName: string): PropertyType {
    if (navigationProperties.has(propertyName)) {
        return PropertyType.navigational;
    } else if (scalarProperties.has(propertyName)) {
        return PropertyType.scalar;
    } else {
        return PropertyType.notProperty;
    }
}

console.log(navigationProperties.size);
console.log(scalarProperties.size);

navigationProperties.forEach(property => {
    if(scalarProperties.has(property)) {
        console.log(property);
    }
})