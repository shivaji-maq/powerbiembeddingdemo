/* eslint-disable @typescript-eslint/no-explicit-any */
import { describe, it, beforeEach, afterEach, vi, expect } from "vitest";
import { render, cleanup } from "@testing-library/react";
import { ThemeWatcher } from "../ThemeWatcher";

// Mock react-redux
vi.mock("react-redux", () => ({
    useSelector: vi.fn(),
}));

import { useSelector } from "react-redux";

describe("ThemeWatcher", () => {
    let matchMediaMock: any;
    let addEventListenerMock: any;
    let removeEventListenerMock: any;
    let originalMatchMedia: any;

    beforeEach(() => {
        cleanup();
        addEventListenerMock = vi.fn();
        removeEventListenerMock = vi.fn();
        matchMediaMock = vi.fn().mockImplementation(query => ({
            matches: false,
            media: query,
            addEventListener: addEventListenerMock,
            removeEventListener: removeEventListenerMock,
        }));
        originalMatchMedia = window.matchMedia;
        Object.defineProperty(window, "matchMedia", {
            writable: true,
            value: matchMediaMock,
        });
        document.documentElement.dataset.theme = "";
    });

    afterEach(() => {
        cleanup();
        window.matchMedia = originalMatchMedia;
        document.documentElement.dataset.theme = "";
        vi.clearAllMocks();
    });

    it("sets theme to light when theme is 'light'", () => {
        (useSelector as any).mockImplementation(() => "light");
        render(<ThemeWatcher />);
        expect(document.documentElement.dataset.theme).toBe("light");
    });

    it("sets theme to dark when theme is 'dark'", () => {
        (useSelector as any).mockImplementation(() => "dark");
        render(<ThemeWatcher />);
        expect(document.documentElement.dataset.theme).toBe("dark");
    });

    it("sets theme to dark when theme is 'system' and prefers-color-scheme is dark", () => {
        matchMediaMock.mockImplementationOnce(() => ({
            matches: true,
            media: "(prefers-color-scheme: dark)",
            addEventListener: addEventListenerMock,
            removeEventListener: removeEventListenerMock,
        }));
        (useSelector as any).mockImplementation(() => "system");
        render(<ThemeWatcher />);
        expect(document.documentElement.dataset.theme).toBe("dark");
    });

    it("sets theme to light when theme is 'system' and prefers-color-scheme is light", () => {
        matchMediaMock.mockImplementationOnce(() => ({
            matches: false,
            media: "(prefers-color-scheme: dark)",
            addEventListener: addEventListenerMock,
            removeEventListener: removeEventListenerMock,
        }));
        (useSelector as any).mockImplementation(() => "system");
        render(<ThemeWatcher />);
        expect(document.documentElement.dataset.theme).toBe("light");
    });

    it("adds and removes event listener when theme is 'system'", () => {
        (useSelector as any).mockImplementation(() => "system");
        render(<ThemeWatcher />);
        expect(addEventListenerMock).toHaveBeenCalledWith("change", expect.any(Function));
        cleanup();
        // After cleanup, removeEventListener should be called
        expect(removeEventListenerMock).toHaveBeenCalledWith("change", expect.any(Function));
    });

    it("renders with data-testid", () => {
        (useSelector as any).mockImplementation(() => "light");
        const { getByTestId } = render(<ThemeWatcher />);
        expect(getByTestId("theme-watcher")).toBeInTheDocument();
    });
});
