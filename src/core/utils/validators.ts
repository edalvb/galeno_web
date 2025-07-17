export class Validators {

    static required(value: string | undefined): string | null {
        if (value === undefined) {
            return "This field is required";
        }
        if (value?.trim() === "") {
            return "This field is required";
        }
        return null;
    }

    static email(value: string | undefined): string | null {
        if (!value) return null;
        if (value?.trim() === "") {
            return "This field is required";
        }
        const emailRegex = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}$/;
        if (!emailRegex.test(value)) {
            return "Please enter a valid email address";
        }
        return null;
    }

    static numeric(value: string | undefined): string | null {
        if (!value) return "This field is required"; // If field is not required and empty
        if (value?.trim() === "") {
            return "This field is required";
        }
        const numericRegex = /^\d+$/;
        if (!numericRegex.test(value)) {
            return "This field only accepts numbers";
        }
        return null;
    }

    static numericWithDecimals(value: string | undefined): string | null {
        if (!value) return null; // If field is not required and empty
        if (value?.trim() === "") {
            return "This field is required";
        }
        const decimalRegex = /^\d*\.?\d+$/;
        if (!decimalRegex.test(value)) {
            return "This field only accepts numbers and decimals";
        }
        return null;
    }

    static phone(value: string | undefined): string | null {
        if (!value) return null; // If field is not required and empty
        if (value?.trim() === "") {
            return "This field is required";
        }
        const phoneRegex = /^\+?[\d\s-]{9,}$/;
        if (!phoneRegex.test(value)) {
            return "Please enter a valid phone number";
        }
        return null;
    }

    // Utility method to combine multiple validators
    static compose(validators: ((value: string | undefined) => string | null)[]): (value: string | undefined) => string | null {
        return (value: string | undefined): string | null => {
            for (const validator of validators) {
                const result = validator(value);
                if (result) return result;
            }
            return null;
        };
    }
}