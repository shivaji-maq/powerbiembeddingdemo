import { createSlice, type PayloadAction } from "@reduxjs/toolkit";
import type { BookmarkSliceState, BookmarkedVisual } from "./@types";

const STORAGE_KEY = "qvc_bookmarks";

function loadBookmarks(): BookmarkedVisual[] {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (!stored) return [];
    const parsed = JSON.parse(stored);
    return Array.isArray(parsed) ? (parsed as BookmarkedVisual[]) : [];
  } catch {
    return [];
  }
}

function saveBookmarks(bookmarks: BookmarkedVisual[]) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(bookmarks));
  } catch { /* storage full or unavailable */ }
}

const initialState: BookmarkSliceState = {
  bookmarks: loadBookmarks(),
};

const bookmarkSlice = createSlice({
  name: "bookmarks",
  initialState,
  reducers: {
    addBookmark(state, action: PayloadAction<BookmarkedVisual>) {
      state.bookmarks.push(action.payload);
      saveBookmarks(state.bookmarks);
    },
    removeBookmark(state, action: PayloadAction<string>) {
      state.bookmarks = state.bookmarks.filter((b) => b.id !== action.payload);
      saveBookmarks(state.bookmarks);
    },
    renameBookmark(state, action: PayloadAction<{ id: string; name: string }>) {
      const bookmark = state.bookmarks.find((b) => b.id === action.payload.id);
      if (bookmark) {
        bookmark.name = action.payload.name;
        saveBookmarks(state.bookmarks);
      }
    },
  },
});

export const { addBookmark, removeBookmark, renameBookmark } = bookmarkSlice.actions;
export default bookmarkSlice.reducer;
