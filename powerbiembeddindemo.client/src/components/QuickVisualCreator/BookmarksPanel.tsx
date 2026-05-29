import { useSelector, useDispatch } from "react-redux";
import type { RootState } from "../../redux/store";
import { removeBookmark } from "../../redux/slices/bookmarkSlice/bookmarkSlice";
import type { BookmarkedVisual } from "../../redux/slices/bookmarkSlice/@types";
import "./BookmarksPanel.css";

interface BookmarksPanelProps {
  onLoadBookmark: (bookmark: BookmarkedVisual) => void;
}

export default function BookmarksPanel({ onLoadBookmark }: BookmarksPanelProps) {
  const dispatch = useDispatch();
  const bookmarks = useSelector((state: RootState) => state.bookmarks.bookmarks);

  if (bookmarks.length === 0) {
    return (
      <div className="bookmarks-panel">
        <h3 className="bookmarks-title">Bookmarked Visuals</h3>
        <p className="bookmarks-empty">No bookmarks yet. Create or update a visual to store it here.</p>
      </div>
    );
  }

  return (
    <div className="bookmarks-panel">
      <h3 className="bookmarks-title">Bookmarked Visuals</h3>
      <div className="bookmarks-list">
        {bookmarks.map((bookmark) => (
          <div key={bookmark.id} className="bookmark-card">
            <div className="bookmark-card-header">
              <span className="bookmark-name">{bookmark.name}</span>
              <span className="bookmark-type">{bookmark.visuals.length} visual{bookmark.visuals.length !== 1 ? "s" : ""}</span>
            </div>
            <div className="bookmark-card-fields">
              {bookmark.visuals.map((v, i) => (
                <span key={i} className="bookmark-field-tag">{v.visualType}</span>
              ))}
            </div>
            <div className="bookmark-card-meta">
              <span className="bookmark-date">
                {new Date(bookmark.createdAt).toLocaleDateString()}
              </span>
            </div>
            <div className="bookmark-card-actions">
              <button
                className="bookmark-load-btn"
                onClick={() => onLoadBookmark(bookmark)}
                title="Load this visual"
              >
                Load
              </button>
              <button
                className="bookmark-delete-btn"
                onClick={() => dispatch(removeBookmark(bookmark.id))}
                title="Remove bookmark"
              >
                Remove
              </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
