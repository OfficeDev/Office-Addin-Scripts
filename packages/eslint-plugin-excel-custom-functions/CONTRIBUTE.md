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

#### Rule Files
```
src\rules\no-office-read-calls.ts
src\rules\no-office-write-calls.ts
```

#### Test Files
```
tests\rules\no-office-read-calls.test.ts
tests\rules\no-office-write-calls.test.ts
```

#### Utils
```
src\rules\utils.ts
```