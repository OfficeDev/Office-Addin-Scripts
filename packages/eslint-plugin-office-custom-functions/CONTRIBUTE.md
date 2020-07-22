To get the plugin up and running for contribution on your machine, follow these steps

## Step 1
Clone Office-Addin-Scripts and navigate to this directory

## Step 2

```
npm i
```

## Step 3

```
npm run build
```

## Step 4

```
npm run test
```

### The files of note are the following:
```
src\rules\no-office-api-calls.ts
tests\rules\no-office-api-calls.test.ts
```

#### The below file is also notable, but only the bottom functions. There is cleanup that's needed.
```
src\rules\utils.ts
```